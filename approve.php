<?php
/**
 * api/approve.php
 *
 * The real payoff of the whole pipeline: takes a 'submitted' job's
 * reviewed extraction items and writes them as real rows in rt_units
 * and rt_lessons, under the job's product_id + grade. This is the only
 * endpoint in the app that writes to the main curriculum tables rather
 * than the ingestion_* tables.
 *
 * Enforces a soft two-person rule from the engineering addendum: the
 * approver cannot be the same person who submitted the job. This is
 * client-supplied and trivially bypassable until real auth exists —
 * it's here so the shape of the rule is in place, not as real security.
 *
 * POST /api/approve.php
 * Body: { "job_id": "uuid", "approved_by": "uuid", "review_notes": "..." }
 *
 * Deleted items (was_deleted=1) are skipped entirely. A lesson whose
 * parent unit was deleted is also skipped (there's nothing to attach it
 * to) and reported back in "skipped_orphaned_lessons" so it's visible,
 * not silently dropped.
 */

require_once __DIR__ . '/../config.php';

header('Content-Type: application/json');
header('Cache-Control: no-store, no-cache, must-revalidate, max-age=0');
header('Pragma: no-cache');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    header('Allow: POST');
    echo json_encode(['error' => 'Method not allowed']);
    exit;
}

try {
    $pdo = get_db_connection();
    $body = json_decode(file_get_contents('php://input'), true);

    $jobId       = $body['job_id'] ?? null;
    $approvedBy  = $body['approved_by'] ?? null;
    $reviewNotes = $body['review_notes'] ?? null;

    $missing = [];
    if (empty($jobId))      $missing[] = 'job_id';
    if (empty($approvedBy)) $missing[] = 'approved_by';
    if (!empty($missing)) {
        http_response_code(400);
        echo json_encode(['error' => 'Missing required field(s)', 'fields' => $missing]);
        exit;
    }

    $stmt = $pdo->prepare('SELECT * FROM rt_ingestion_jobs WHERE id = :id');
    $stmt->execute(['id' => $jobId]);
    $job = $stmt->fetch();

    if (!$job) {
        http_response_code(404);
        echo json_encode(['error' => 'Job not found']);
        exit;
    }

    if ($job['status'] !== 'submitted') {
        http_response_code(409);
        echo json_encode(['error' => "Job is at status '{$job['status']}' — only 'submitted' jobs can be approved."]);
        exit;
    }

    if (empty($job['product_id']) || empty($job['grade'])) {
        http_response_code(400);
        echo json_encode(['error' => 'Job is missing product_id and/or grade — cannot publish without both.']);
        exit;
    }

    // Soft two-person rule: approver != submitter.
    if ($approvedBy === $job['submitted_by']) {
        http_response_code(403);
        echo json_encode([
            'error' => 'The approver must be different from the person who submitted the job.',
        ]);
        exit;
    }

    $stmt = $pdo->prepare('SELECT id FROM rt_users WHERE id = :id');
    $stmt->execute(['id' => $approvedBy]);
    if (!$stmt->fetch()) {
        http_response_code(400);
        echo json_encode(['error' => 'approved_by does not exist']);
        exit;
    }

    // ---- Load non-deleted items in reading order ----
    $stmt = $pdo->prepare(
        'SELECT * FROM rt_ingestion_extraction_items
         WHERE job_id = :job_id AND was_deleted = FALSE
         ORDER BY sequence'
    );
    $stmt->execute(['job_id' => $jobId]);
    $items = $stmt->fetchAll();

    $pdo->beginTransaction();
    try {
        // Starting unit_number for this product+grade — units created in
        // this job continue numbering from whatever already exists,
        // rather than always starting at 1.
        $stmt = $pdo->prepare(
            'SELECT COALESCE(MAX(unit_number), 0) FROM rt_units WHERE product_id = :product_id AND grade = :grade'
        );
        $stmt->execute(['product_id' => $job['product_id'], 'grade' => $job['grade']]);
        $nextUnitNumber = (int) $stmt->fetchColumn() + 1;

        $insertUnit = $pdo->prepare(
            'INSERT INTO rt_units (id, product_id, grade, unit_number, title, total_lessons, confidence)
             VALUES (:id, :product_id, :grade, :unit_number, :title, :total_lessons, :confidence)'
        );
        $insertLesson = $pdo->prepare(
            'INSERT INTO rt_lessons (id, unit_id, lesson_number, title, source_page_start, source_page_end, confidence)
             VALUES (:id, :unit_id, :lesson_number, :title, :source_page_start, :source_page_end, :confidence)'
        );

        $unitIdBySequence = [];
        $unitsCreated = 0;
        $lessonsCreated = 0;
        $skippedOrphanedLessons = [];

        // Pass 1: units. Need total_lessons per unit before inserting, so
        // count lesson children (non-deleted) up front per unit sequence.
        $lessonCountByParentSeq = [];
        foreach ($items as $item) {
            if ($item['item_type'] === 'lesson' && $item['parent_sequence'] !== null) {
                $lessonCountByParentSeq[$item['parent_sequence']] = ($lessonCountByParentSeq[$item['parent_sequence']] ?? 0) + 1;
            }
        }

        foreach ($items as $item) {
            if ($item['item_type'] !== 'unit') {
                continue;
            }
            $unitId = generate_uuid_v4();
            $insertUnit->execute([
                'id'            => $unitId,
                'product_id'    => $job['product_id'],
                'grade'         => $job['grade'],
                'unit_number'   => $nextUnitNumber,
                'title'         => $item['accepted_title'] ?? $item['proposed_title'],
                'total_lessons' => $lessonCountByParentSeq[$item['sequence']] ?? 0,
                'confidence'    => $item['confidence'],
            ]);
            $unitIdBySequence[$item['sequence']] = $unitId;
            $nextUnitNumber++;
            $unitsCreated++;
        }

        // Pass 2: lessons, now that every unit's id is known.
        $lessonNumberByUnitId = [];
        $createdLessons = []; // full records, for the slide-deck webhook below
        foreach ($items as $item) {
            if ($item['item_type'] !== 'lesson') {
                continue;
            }
            $parentUnitId = $unitIdBySequence[$item['parent_sequence']] ?? null;
            if ($parentUnitId === null) {
                $skippedOrphanedLessons[] = ['sequence' => $item['sequence'], 'proposed_title' => $item['proposed_title']];
                continue;
            }
            $lessonNumberByUnitId[$parentUnitId] = ($lessonNumberByUnitId[$parentUnitId] ?? 0) + 1;

            $lessonId = generate_uuid_v4();
            $insertLesson->execute([
                'id'                => $lessonId,
                'unit_id'           => $parentUnitId,
                'lesson_number'     => $lessonNumberByUnitId[$parentUnitId],
                'title'             => $item['accepted_title'] ?? $item['proposed_title'],
                'source_page_start' => $item['source_page_start'],
                'source_page_end'   => $item['source_page_end'],
                'confidence'        => $item['confidence'],
            ]);
            $lessonsCreated++;
            $createdLessons[] = [
                'lesson_id'         => $lessonId,
                'unit_id'           => $parentUnitId,
                'lesson_number'     => $lessonNumberByUnitId[$parentUnitId],
                'title'             => $item['accepted_title'] ?? $item['proposed_title'],
                'source_page_start' => $item['source_page_start'],
                'source_page_end'   => $item['source_page_end'],
            ];
        }

        // Map unit id -> title, so the webhook payload can include the
        // unit each lesson belongs to without a second query.
        $unitTitleById = [];
        foreach ($items as $item) {
            if ($item['item_type'] === 'unit' && isset($unitIdBySequence[$item['sequence']])) {
                $unitTitleById[$unitIdBySequence[$item['sequence']]] = $item['accepted_title'] ?? $item['proposed_title'];
            }
        }

        $stmt = $pdo->prepare(
            'UPDATE rt_ingestion_jobs
             SET status = \'approved\', reviewed_by = :reviewed_by, reviewed_at = NOW(),
                 review_notes = :notes, published_at = NOW()
             WHERE id = :id'
        );
        $stmt->execute(['reviewed_by' => $approvedBy, 'notes' => $reviewNotes, 'id' => $jobId]);

        $pdo->commit();
    } catch (Throwable $e) {
        $pdo->rollBack();
        throw $e;
    }

    // ---- Notify the slide-deck service, once per newly-created lesson ----
    // Deliberately AFTER commit and wrapped so a slow/unreachable slide
    // service can never fail or roll back a real approval. Best-effort
    // only -- if this fails, the lesson still exists in rt_lessons and
    // slide generation can be retried/triggered manually later.
    //
    // KNOWN GAP: lesson_text is not yet populated -- rt_ingestion_extraction_items
    // does not carry the lesson's actual prose, only title/page-range
    // metadata. Until that's resolved, the slide-deck service receives a
    // page reference it cannot yet act on. See webhook payload below.
    foreach ($createdLessons as $lesson) {
        notify_slide_deck_service([
            'job_id'            => $jobId,
            'lesson_id'         => $lesson['lesson_id'],
            'unit_id'           => $lesson['unit_id'],
            'unit_title'        => $unitTitleById[$lesson['unit_id']] ?? null,
            'lesson_title'      => $lesson['title'],
            'lesson_number'     => $lesson['lesson_number'],
            'product_id'        => $job['product_id'],
            'grade'             => $job['grade'],
            'source_page_start' => $lesson['source_page_start'],
            'source_page_end'   => $lesson['source_page_end'],
            'lesson_text'       => null, // TODO: unresolved -- see note above
        ]);
    }

    $stmt = $pdo->prepare('SELECT * FROM rt_ingestion_jobs WHERE id = :id');
    $stmt->execute(['id' => $jobId]);

    echo json_encode([
        'job' => $stmt->fetch(),
        'units_created' => $unitsCreated,
        'lessons_created' => $lessonsCreated,
        'skipped_orphaned_lessons' => $skippedOrphanedLessons,
    ]);

} catch (Throwable $e) {
    http_response_code(500);
    echo json_encode(['error' => 'Unexpected error — job was NOT published (transaction rolled back)', 'detail' => $e->getMessage()]);
}

function generate_uuid_v4(): string
{
    $data = random_bytes(16);
    $data[6] = chr(ord($data[6]) & 0x0f | 0x40);
    $data[8] = chr(ord($data[8]) & 0x3f | 0x80);
    return vsprintf('%s%s-%s-%s-%s-%s%s%s', str_split(bin2hex($data), 4));
}

/**
 * Best-effort notification to the slide-deck generation service. Uses a
 * short connect/response timeout and swallows all failures -- this must
 * never be able to slow down or break a real approval. The slide-deck
 * service is expected to respond immediately (202-style) and do the
 * actual extraction/generation work asynchronously on its own side, not
 * make this request wait for a full deck to be built.
 *
 * Requires SLIDE_DECK_SERVICE_URL and SLIDE_DECK_SERVICE_TOKEN to be
 * defined in config.php. If either is missing, this is a no-op --
 * slide-deck integration is optional, not a hard dependency of approval.
 */
function notify_slide_deck_service(array $payload): void
{
    if (!defined('SLIDE_DECK_SERVICE_URL') || !defined('SLIDE_DECK_SERVICE_TOKEN')) {
        return;
    }
    if (empty(SLIDE_DECK_SERVICE_URL)) {
        return;
    }

    try {
        $ch = curl_init(rtrim(SLIDE_DECK_SERVICE_URL, '/') . '/webhook/lesson-approved');
        curl_setopt_array($ch, [
            CURLOPT_POST           => true,
            CURLOPT_POSTFIELDS     => json_encode($payload),
            CURLOPT_HTTPHEADER     => [
                'Content-Type: application/json',
                'Authorization: Bearer ' . SLIDE_DECK_SERVICE_TOKEN,
            ],
            CURLOPT_CONNECTTIMEOUT => 2,
            CURLOPT_TIMEOUT        => 3,
            CURLOPT_RETURNTRANSFER => true,
        ]);
        curl_exec($ch);
        if (curl_errno($ch)) {
            error_log('slide-deck webhook failed (non-fatal): ' . curl_error($ch)
                . ' for lesson_id=' . ($payload['lesson_id'] ?? 'unknown'));
        }
        curl_close($ch);
    } catch (Throwable $e) {
        // Never let a notification failure propagate -- approval already
        // committed successfully by the time this runs.
        error_log('slide-deck webhook exception (non-fatal): ' . $e->getMessage());
    }
}

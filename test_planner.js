const pptxgen = require("pptxgenjs");
const { renderPlannerSections } = require("./planner_renderer.js");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

function titleSlide(title) {
  const s = pres.addSlide();
  s.addText(title, { x: 0.55, y: 0.35, w: 12.2, h: 0.6, fontFace: "Arial", fontSize: 28, bold: true, color: "382DB0", margin: 0 });
  return s;
}

// ===== Test 1: Realistic Narrative Writing Planner (Lesson 18b) =====
// Real content from RT_Gr4U2_KNO_full_unit_Teacher_Guide.pdf, page ~389
{
  const s = titleSlide("Test 1: Realistic Narrative Writing Planner");
  const sections = [
    { type: "text", label: "Message to Readers", content: "Sometimes changes that feel devastating and horrible at first turn out to be for the best and make us happier in the long run." },
    { type: "list", label: "Protagonist", items: ["Just finished Grade 4.", "Only child.", "Close with parents."] },
    { type: "list", label: "Change Protagonist Faces", items: ["Parents getting divorced.", "Family won't live together anymore."] },
    { type: "table", label: "Conflict and Sequence of Events", columns: ["Introduction of Conflict", "Rising Action", "Resolution"], rows: [
      [
        ["Parents tell protagonist they're getting divorced.", "Protagonist reacts by running away.", "Thinks they are losing both parents.", "Devastated."],
        ["Later in time after family separated.", "Activities no longer fun.", "Contrast to memories of when parents together and having family fun."],
        ["Pivotal moment\u2014mom and dad working together again. Enjoy a visit.", "Realizes it had actually been a long time since family was happy.", "Parents happier and everyone can have fun together and feel better after divorce.", "Sees they have both parents and life is actually better."]
      ]
    ] },
  ];
  renderPlannerSections(s, sections, 0.55, 1.1, 12.2, 6.0);
}

// ===== Test 2: Historical Figure Fact Card Planner (Lesson 4b) =====
{
  const s = titleSlide("Test 2: Historical Figure Fact Card Planner");
  const sections = [
    { type: "table", columns: ["Research Question", "Notes"], rows: [
      ["What is your figure's background?", ""],
      ["What are your figure's most important accomplishments?", ""],
      ["What is an interesting fact about your figure?", ""],
      ["Which of your figure's quotes best represents their beliefs or experience?", ""],
      ["What is your figure's greatest contribution to the Great Migration?", ""],
      ["Additional Research Questions", ""],
    ] },
  ];
  renderPlannerSections(s, sections, 0.55, 1.1, 12.2, 6.0);
}

// ===== Test 3: Figurative Language reference chart (Lesson 12b) =====
{
  const s = titleSlide("Test 3: Figurative Language Chart");
  const sections = [
    { type: "table", label: "Figurative Language", columns: ["Term", "Definition", "Examples"], rows: [
      ["simile", "A comparison of one thing to another using \u201clike\u201d or \u201cas.\u201d", ["Her hug was as warm as the sun.", "The puppy's teeth were like knives.", "He's as fast as a cheetah."]],
      ["metaphor", "A direct comparison of two seemingly different things to convey an idea; it does not need to use the word \u201clike\u201d or \u201cas.\u201d", ["She's a busy bee!", "The autumn leaves were colorful confetti.", "The racer is a cheetah."]],
      ["hyperbole", "A statement or claim that is intentionally exaggerated or overstated for emphasis or effect.", ["I have a mountain of laundry to fold.", "I am so hungry I could eat a horse."]],
      ["personification", "Giving human characteristics or behaviors to nonhuman objects or living things.", ["The wind grabbed my hat and threw it into the air.", "The dolphins danced gracefully in the water."]],
    ] },
  ];
  renderPlannerSections(s, sections, 0.55, 1.1, 12.2, 6.0);
}

pres.writeFile({ fileName: "planner_prototype.pptx" }).then(() => console.log("wrote planner_prototype.pptx"));

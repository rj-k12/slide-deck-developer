# replit.nix -- installs Python, Node, and LibreOffice together.
# Same three runtime dependencies as every other deployment option
# (see DEPLOY.md / DEPLOY_NO_DOCKER.md) -- Replit's Nix package system
# just expresses it differently.

{ pkgs }: {
  deps = [
    pkgs.nodejs_20
    pkgs.python311
    pkgs.python311Packages.pip
    pkgs.libreoffice-still   # heavy (~600MB+) -- see note in DEPLOY_REPLIT.md
    pkgs.poppler_utils       # pdftoppm, used by make_cover.py
  ];
}

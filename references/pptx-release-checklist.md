# PowerPoint Release Checklist

## Before mutation
- Identify the requested artifact and exact source revision. Do not use an older installed skill when the newer supplied package is available.
- Inventory text, shape IDs, notes, links, media, theme/font declarations, and layout assumptions.
- Decide which changes are mechanical and which require semantic review. Choose Inter only under the explicit Inter profile, or the legacy Arial presentation profile when requested.
- Keep the source file and original packaged Skill immutable.

## Brand treatment
- Use original logo/Grid artwork, aspect-preserving placement, visible-ink minimum width, and real clear space.
- Preserve approved artwork bytes; do not colour-map embedded graphics or invent product identities.
- Use palette tokens, including separate OT Grey and the earlier neutral greys. No palette strip belongs on an ordinary customer slide.
- Keep title, body, tables, labels, charts, captions, master/layout defaults, and notes consistent where relevant. Preserve symbol/CJK/complex-script declarations when unsupported by Inter.
- Do not fix dense content by silently shrinking the entire slide. Split content or change the layout with explicit approval.

## Structural and rendered validation
- Confirm source text and hyperlinks are preserved; include numeric/chart data checks where applicable.
- Confirm unaffected opaque parts/media are unchanged. Check per-instance image operations and source metadata.
- Confirm all local relationships resolve and ZIP integrity passes. Do not follow external relationships.
- Render every slide. Inspect the rendered fonts, required labels, aspect ratio, and clipping/overlap.
- Document intentional containments/connector intersections rather than disabling all overlap checks.
- Inspect all output PNGs at useful scale; retain the result separately from automated test output.

## Release boundary
- A structural audit is not visual certification or Fortinet brand-team approval.
- A Linux LibreOffice render is not proof of PowerPoint/Keynote parity on Windows/macOS.
- Do not package fonts, customer incident records, credentials, raw working directories, full unapproved libraries, or temporary diagnostics.
- Use versioned provenance and file hashes. Keep the Skill ZIP at or below 25 MiB.
- State exactly what was packaged, tested, installed, or published. Do not conflate them.

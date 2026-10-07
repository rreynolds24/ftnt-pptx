# PowerPoint Build Learnings and Design Rationale

This is a reusable engineering summary of the completed slide build, recovered build records, and the new offline regression results. It is not a transcript of internal deliberation.

## Preserve the intent before applying a visual system

Inventory the supplied slide's title, section headings, all product labels, environmental nodes, central message, and roadmap stages. A brand refresh must not silently remove a difficult label or introduce a capability claim. For the completed example, six capability sections, eight environmental nodes, and four roadmap stages were the useful invariants. Those counts belong to this template, not every Fortinet slide.

Represent a conceptual architecture with native text, containers, and editable connector shapes. Keep the artwork in individual image objects. This allows future edits without regenerating the entire illustration. Do not describe raster artwork as editable vector paths. Do not claim an operational topology, traffic direction, or supported integration merely because two concepts are connected on a slide.

## Use source authority rather than a resemblance

Use the official corporate logo and Grid from the supplied kit. A typed wordmark or a shield with a graduation cap is not a substitute. Use the official guideline's generic or pillar icon when it genuinely represents the concept, and label that role. A generic AI icon must not be described as an official FortiAIGate product logo. An unverified online library can be staged for review, not silently certified as current.

Inter was explicitly requested for this work. Keep that approved presentation profile while retaining the 2024 guide's separate Arial exception. A local preference does not rewrite the historical guide. Use exact palette tokens from the source, including approved tints and neutrals; keep sealed official artwork untouched even when its source rendering contains a near-match colour.

## Keep official artwork intact and measurable

Use aspect-preserving contain placement. Do not trim the logo's transparent margins just to fill a box. Those margins do not remove the need to check cap-height clear space around the visible ink. Measure the visible artwork, not only the outer image canvas, against the minimum width. Check the logo and Grid colours against the actual background and leave the area free of unrelated text, diagram edges, and decoration.

When the official standalone icon is unavailable but the supplied brand guide contains the required artwork, an exact source crop is a defensible fallback. Record PDF hash, page, rectangle in PDF points, output dimensions, and asset hash. The original crop-to-SVG approach produced disproportionately large SVGs because clipped views can retain whole-page data and hidden artwork. Prefer compact PNG crops at an appropriate display resolution in that case. Do not call them a native official icon library.

## Layout lessons that survived rendering

The completed slide uses a white canvas, a restrained red accent, neutral panels, two three-card columns, and a conceptual hub between eight node cards. That structure replaced a busy illustrated campus scene and made the content editable. It is one useful pattern, not an official mandatory layout.

Reserve title, subtitle, body, and roadmap bands before placing content. Named objects and a layout manifest make subsequent changes deterministic. Place connection buses behind the cards so the diagram does not cross over labels. Keep peer panel spacing consistent.

The first render revealed clipped/wrapped card copy despite apparently valid box geometry. Increasing the relevant row heights and panel heights resolved this more reliably than indiscriminate font shrinking. A long heading needed local treatment. Captured PDF text bounding boxes helped locate the actual wraps, but visual inspection remained necessary.

Roadmap construction needed special care. Rotating a triangle or using an inappropriate arrow shape changes the effective bounding geometry. A native chevron plus a small same-colour patch at the opening edge worked, after keeping that patch outside the title text. Document that intentional overlap; do not suppress all overlap warnings. Diagnostic passes that exclude containers must not also hide genuine text-on-text collisions.

The final example contains small node and roadmap text. Preserve that observation as a density warning. It is not evidence that 9-11 pt text is generally appropriate for presentation screens. Use the template for a dense downloadable explainer or split it into multiple slides for projection.

## Validation is three different questions

1. Structural: Is the package valid? Are text, sources, shape counts, notes, hyperlinks, and unaffected binaries preserved? Are dimensions and relationships sensible?
2. Rendered: Did the output actually use Inter? Are all required labels visible? Is the slide count and aspect ratio correct? Did the renderer substitute a font?
3. Human visual: Are the hierarchy, whitespace, line wrapping, logo clear space, diagram routing, contrast, and density acceptable?

Passing one question does not answer the others. A no-overflow result is not proof of readability, correct icon meaning, official brand approval, or full Office compatibility. PDF text checks are useful for omissions but cannot determine semantic completeness on their own.

## Deterministic operations versus judgement

Automate palette lookup, catalog search, font declarations, reviewed colour maps, explicit shape/table styling, per-instance picture replacement, hash verification, package preservation, and rendering reports. Keep content interpretation, product identity, new topology, automatic layout redesign, and final visual approval outside the blind transformation path.

For an existing deck, audit first, write a hash-bound plan, and apply to a new file. Do not edit shared media globally when replacing one icon, do not execute external links, and do not replace the original file. Preserve unknown OOXML parts and namespace prefixes rather than round-tripping the whole deck through a lossy authoring model.

## Reuse without prompt replay

Use `assets/brand-tokens.json` and `assets/asset-catalog.json` as the machine-readable control plane. Use the offline CLI before loading long reference texts. Read detailed source guidance only for exceptions. Use the JSON-driven builder for this repeatable explainer layout instead of regenerating its positioning code from a long prompt.

The original working source files are not mounted in this session. The reusable crop recipe, architecture pattern, iterative fixes, and validation approach were recovered from the build record and completed artifact. Delivered scripts are new, parameterised implementations of those behaviours, with fresh tests. No claim is made that they are byte-identical copies of the previous temporary scripts.

## Additional crop-boundary finding during this release

Catalog preview exposed a clipped right edge on the previously recovered cloud icon. The source guide showed the full edge alongside an unrelated connector. The crop was widened and only the separate lower-right connector region was excluded. The recipe records both regions and the catalog retains the original recovered hash. Do not blindly promote an earlier crop as complete because it came from a successful slide build; inspect each asset independently at useful scale. This correction changes isolation bounds, not the underlying logo/icon design.

# Offline PowerPoint Toolkit

## Contents
1. Runtime and scope
2. Audit, plan, and apply
3. Explicit shape, table, and image operations
4. Local asset operations
5. Rendering and validation
6. Reusable architecture builder
7. Limitations and safety

## Runtime and scope

`fortinet_pptx.py` uses Python 3.10+ and the standard library only. It does not use an API key, model, network connection, LibreOffice, or an installed Skill to inspect and modify a supported existing `.pptx` file.

The accompanying `assets/` files are runtime data, not a remote dependency. Keep the toolkit directory structure intact. The full Skill also contains editorial authorities and visual references; the smaller standalone toolkit contains only the scripts, palette, verified logo/icon catalog, examples, and operating instructions needed by these utilities.

Tested here on Linux with Python 3.13.5, PptxGenJS 4.0.0, LibreOffice 25.2.3.2, and PyMuPDF 1.26.7. The Python core uses cross-platform standard-library APIs. Windows/macOS execution and Microsoft PowerPoint/Keynote rendering were not tested in this release.

## Audit, plan, and apply

Run from the unpacked toolkit or Skill root:

```sh
python scripts/fortinet_pptx.py audit input.pptx --output inventory.json
python scripts/fortinet_pptx.py plan input.pptx --font Inter --theme-colours --map-colour EE3124=DA291C --output plan.json
python scripts/fortinet_pptx.py apply input.pptx --plan plan.json --output branded.pptx --report branded.audit.json
```

On Windows, `py -3` can replace `python`. The wrapper `scripts/brand-pptx.cmd` also forwards arguments to the Python core. On POSIX systems, `scripts/brand-pptx.sh` uses `python3`.

The audit lists slides in presentation order, actual slide-part paths, shape IDs, names, text, explicit dimensions, declared Latin fonts, direct RGB values, existing picture hashes, and review findings. It lists external relationships but never retrieves them.

The default plan changes the Latin text-font profile to Inter. Theme recolouring and direct RGB remapping are opt-in. Theme recolouring can alter inherited chart, table, or semantic colours, so inspect the plan and the rendered output. The example red remapping is explicit, not a general nearest-colour classifier.

A plan is bound to the exact input deck, palette file, and catalog hashes. A changed input or data file requires a new plan. Editing intended operations in the JSON is supported; applying the plan is the explicit mutation step. The source file is never replaced. Existing destination files are protected unless `--force` is used; `--force` still cannot overwrite the source path.

## Explicit shape, table, and image operations

Start with the generated plan. Keep its hashes and standard fields. Copy only the relevant operations into the existing arrays, using IDs from the audit of that specific deck.

```json
{
  "style_shapes": [
    {
      "slide_part": "ppt/slides/slide1.xml",
      "shape_id": 2,
      "font_size_pt": 28,
      "font_colour": "000000",
      "bold": true,
      "align": "left"
    }
  ],
  "table_styles": [
    {
      "slide_part": "ppt/slides/slide1.xml",
      "shape_id": 5,
      "header_fill": "DA291C",
      "header_text": "FFFFFF",
      "body_fill": "FFFFFF",
      "body_text": "000000"
    }
  ]
}
```

Shape operations also accept `fill_colour`, `line_colour`, and `box_inches: [x,y,width,height]`. Geometry is in inches. Geometry changes are explicit, not automatic text-fitting. Grouped geometry and inherited boxes are unsupported. These operations do not rewrite text.

```json
{
  "replace_images": [
    {
      "slide_part": "ppt/slides/slide1.xml",
      "shape_id": 3,
      "asset_id": "secure-networking"
    }
  ],
  "add_logos": [
    {
      "slide_part": "ppt/slides/slide1.xml",
      "asset_id": "fortinet-logo-rgb-black-red",
      "box_inches": [10.5, 0.1, 2.5, 0.75],
      "background_hex": "FFFFFF",
      "clear_space_confirmed": true
    }
  ]
}
```

The example IDs and coordinates are illustrative; they are not universal. Add a logo only in an explicitly reserved area after inspecting the slide, layout, and master. `clear_space_confirmed` records that inspection; it is not an automated collision guarantee. Use `brand_context_confirmed: true` for Grid-only placement after establishing the full brand.

Image replacements act on one selected picture instance. They do not replace a shared media blob used elsewhere. Pictures retain their slot, use aspect-preserving contain placement, and use a verified bundled PNG. Any previous crop and decorative image effects are removed so the new official artwork is not altered. Rotated/flipped/grouped targets are blocked. There is no guessed product-logo matching and no deletion of text pretending to be a logo.

For logo/Grid placement, the tool checks visible-ink minimum width, contrast against the declared background, and the required cap-height-based clear zone staying within the slide. It reports the zone for visual review. It does not prove that a declared background is truly uniform or that inherited objects do not intrude.

## Local asset operations

```sh
python scripts/fortinet_pptx.py icons --query "security"
python scripts/fortinet_pptx.py icons --query "logo" --output logo-options.json
python scripts/brand_assets.py import-drawio local-library.xml --output staged-icons --source-label "Locally supplied library; provenance pending"
```

The catalog contains seven original logo/Grid PNG variants, their original SVG counterparts, and eight source-traceable guideline-derived icon crops. It does not contain the full Fortinet icon library. Search is lexical over IDs, labels, and aliases. It returns choices, not an automatic classification.

The Draw.io importer handles embedded PNG and SVG `data:` entries in a local `<mxlibrary>`. It rejects active/external SVG content and records skipped entries. Compressed shape-XML entries and external URLs are not fetched or guessed. Imported assets are staged with `approved: false`, even when a supplied label calls them official. They require provenance review, visual review, suitable metadata, and a PNG fallback before catalog admission. No imported library is silently merged into production.

To reproduce the source-specific guideline crops:

```sh
python scripts/brand_assets.py extract-pdf FTNT-Brand-Guidelines.pdf --recipe assets/guideline-crop-recipe.json --output regenerated-crops
```

This command requires local PyMuPDF and the exact hash-pinned source PDF. It renders only the specified crop to PNG and records provenance. Do not export a full page into an SVG and call it a small standalone icon; the prior build's vector-crop experiment retained excessive hidden/page content. Raster crops remain labelled as crops, not native official SVGs.

## Rendering and validation

Optional local tools: LibreOffice and PyMuPDF. Install these separately before taking the workflow offline. The toolkit never installs packages or downloads fonts.

```sh
python scripts/render_check.py branded.pptx --output-dir rendered --font Inter
```

This creates PDF, one PNG per slide, a conversion log, and a JSON report. It checks the page count, extracted font names, optional required terms, and structural findings. `--expected-terms required-terms.json` accepts a JSON list of exact text strings.

The check always records `visual_review_required: true`. Inspect every PNG for clipping, crowded lines, logo clear space, contrast, and unintended overlap. Inter declarations in XML alone do not prove the renderer used Inter. Font-name checks alone do not prove visual quality. Do not distribute a file with font fallback as fully brand-compliant.

## Reusable architecture builder

```sh
node scripts/build_connected_architecture.cjs --data examples/connected-architecture.json --output architecture.pptx
```

Node.js 18+ and local PptxGenJS 4.0.0 are required. Dependency information is in `scripts/package.json`; `node_modules` and font files are not included. Resolve dependencies before offline use.

The builder takes six capability cards, eight environmental nodes, a central message, and four roadmap steps. It creates native text, shapes, and non-directional connection buses, embeds verified artwork, writes source notes, and emits a layout manifest. It contains no customer-specific network configuration.

For ChatGPT's slide workflow, set `FORTINET_PPTX_HELPERS` to its supplied Slides helper module so the platform diagnostics and image-fitting helpers run. Outside that environment, a small independent geometry calculation provides contain placement. No host-specific absolute path is embedded in the delivered script.

This is a dense explainer pattern. Several source-derived node/roadmap labels are below 10 pt; that is retained and flagged, not promoted as a general presentation font standard. Split the content for large-room projection. The example is a reproduction of source content, not proof of current product capabilities, and it does not certify a network design.

## Limitations and safety

- Only unencrypted, unsigned, macro-free Transitional OOXML `.pptx` is writable. Embedded-font packages are not repackaged. Audit is read-only.
- Font changes apply to Latin declarations, relevant run defaults, themes, presentation defaults, masters, layouts, tables, charts, and notes. Symbol fonts and existing East Asian/complex-script declarations are preserved. Mixed-script font support still needs review.
- Existing text, source notes, hyperlinks, non-targeted content, and opaque binaries are retained. Changed XML is reserialized with namespace declarations preserved, including prefixes used only in `mc:Ignorable`.
- SmartArt and grouped objects receive applicable font-level treatment, not intelligent reconstruction. Their geometry, animation, and layout behaviour require review. Unknown extension/binary parts remain unchanged.
- A flattened slide image cannot become editable through formatting. No OCR or image classification runs. Use the normal slide-reconstruction workflow instead.
- Theme/direct colour mapping never recolours embedded pictures, logos, or imported brand artwork. Changing a theme can still change semantic chart colours.
- External images, linked workbooks, fonts, or other resources are not downloaded. Local files are never uploaded. Existing relationships are preserved, not executed.
- ZIP/XML inputs are bounded; unsafe paths, duplicate entries, symlinks, DTDs/entities, and invalid operations fail closed.
- Structural preservation is tested; full PowerPoint/Keynote/Windows/macOS visual parity remains unproven.

Render only trusted presentations. Externally linked media/data relationships are rejected before invoking LibreOffice; ordinary hyperlinks are retained and are not followed. The renderer is an optional local application, not a hardened document sandbox.

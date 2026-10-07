# Fortinet Branding 2.0.0

This package combines the supplied Fortinet visual/editorial authorities with a portable offline PowerPoint toolkit. Start with `SKILL.md` in ChatGPT/Codex, or use the CLI directly without an AI session.

```sh
python scripts/fortinet_pptx.py audit input.pptx --output audit.json
python scripts/fortinet_pptx.py plan input.pptx --font Inter --theme-colours --output plan.json
python scripts/fortinet_pptx.py apply input.pptx --plan plan.json --output branded.pptx --report branded.audit.json
```

The source deck is never overwritten. The plan is editable and hash-bound. Font changes can reflow text; theme changes can alter semantic chart colours. Image/logo replacements require explicit asset and shape mappings. Render and inspect every slide before release.

See `references/offline-pptx-toolkit.md` for table/shape styling, logo placement, local library import, PDF crops, optional rendering, dependencies, examples, and limits. See `references/learning/learning-candidates.md` for the promoted learnings and `references/version.md` for source and release scope.

No model calls, credentials, internet access, or font downloads are needed by the core utility. Optional rendering and new-slide generation need local dependencies installed in advance. The bundled assets are a verified subset, not the complete Fortinet icon library.

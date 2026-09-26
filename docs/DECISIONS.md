# Architecture Decision Log

## ADR-008: Admin Forms incorporated from AdminForms repo

**Decision:** Integrate https://github.com/sachchinannam404/AdminForms as a first-class SPO-Hub web part under `src/adminForms/` + `src/webparts/adminForms/`.

**Rationale:** Delivers Request Center, My Requests, Approval Center, multi-type forms, audit, reporting, and Power Automate notifications required by Phase 4 without rewriting a proven suite.

**Integration approach:**
- Source under `src/adminForms/` (models, services, components, config registry)
- Web part: `AdminFormsWebPart` in the SPO-Hub package (group SPO-Hub)
- Power Automate webhook via property pane
- Optional list/group provisioning on load
- Does not replace the generic List Data Adapter used by Tiles/Accordion/etc.

**Consequence:** SPO-Hub ships list-driven content renderers and a full office-admin request suite in one package.

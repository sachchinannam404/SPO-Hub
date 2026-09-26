# SPO-Hub — SharePoint Adoption & Automation Web Parts

Enterprise-grade SharePoint Framework (SPFx) component library focused on SharePoint adoption, employee engagement, knowledge discovery, business-process self-service, and automation.

**Primary principle:** A common, configurable **List-Driven Rendering Framework**. Specialized web parts are reusable renderers on top of that framework.

## Architecture

```
SharePoint List
       │
       ▼
List Data Adapter (SharePointListDataService)
       │
       ▼
Query / Filter / Sort / Paging
       │
       ▼
Normalized Component Model
       │
       ▼
Renderer
   ├── Tiles
   ├── Accordion
   ├── Quotes
   ├── Charts
   ├── Employee Spotlight
   ├── Q&A
   └── Timeline
```

Individual web parts **do not** duplicate SharePoint connection handling, list retrieval, filtering, sorting, audience targeting, error handling, telemetry, localization, or configuration validation.

## Web Part Catalog

| Web Part | Purpose | Status |
|----------|---------|--------|
| **List-Based Tiles** | Application launcher, quick links, automation catalog, department links | Implemented |
| **List-Based Accordion** | FAQ, policies, procedures, knowledge articles | Renderer ready |
| **List-Based Quotes** | Leadership messages, testimonials, adoption messaging | Foundation ready |
| **Employee Spotlight** | Recognition, SME spotlight, new joiners | Foundation ready |
| **List-Based Q&A** | Community questions, process Q&A | Foundation ready |
| **Vertical Timeline** | Roadmaps, milestones, adoption journey | Implemented |
| **List-Based Charts** | Adoption metrics, automation KPIs | Foundation ready |

## Project Structure

```
src/
├── components/
│   ├── common/           # Loading, Empty, Error, Unauthorized states
│   ├── tiles/
│   ├── accordion/
│   ├── quotes/
│   ├── charts/
│   ├── employeeSpotlight/
│   ├── questionsAnswers/
│   └── timeline/
├── services/
│   ├── SharePointListDataService.ts   # Core List Data Adapter
│   ├── PermissionService.ts
│   ├── AudienceService.ts
│   ├── AutomationService.ts
│   ├── TelemetryService.ts
│   └── ErrorService.ts
├── models/
│   ├── ListItem.ts                    # Normalized models + field mapping
│   └── Configuration.ts               # Shared config contracts
├── hooks/
│   └── useListData.ts                 # Shared data-loading hook
├── utils/
└── webparts/
    ├── tiles/
    ├── accordion/
    ├── quotes/
    ├── charts/
    ├── employeeSpotlight/
    ├── questionsAnswers/
    └── timeline/
```

## Getting Started

### Prerequisites

- Node.js `>=18.17.1 <19.0.0` **or** `>=22.14.0 <23.0.0` (SPFx 1.20 / 1.23 supported ranges)
- Gulp CLI, Yeoman, `@microsoft/generator-sharepoint`

```bash
npm install gulp-cli yo @microsoft/generator-sharepoint --global
```

### Install & Build

```bash
npm install
gulp bundle
gulp package-solution
```

The package is produced at `sharepoint/solution/spo-hub.sppkg`.

### Local Workbench / Serve

```bash
gulp serve
```

## Configuration

Every web part exposes a consistent property pane:

**Data Source**
- List Title / List ID
- View
- Field mapping (Title, Description, Image, Link, Category, Order, IsActive, Audience, …)
- Filter, Sort, Item Limit, Refresh Interval

**Display**
- Title, Layout, Columns, Card Size, Image Position, Show Description / Category / Metadata

**Behavior**
- Search, Filter, Pagination, Deep Link, Open in New Tab, Auto Refresh

**Personalization**
- Audience Targeting (Department, Country, Role, M365 Group)

**Automation**
- Power Automate Flow URL, Action Label, Confirmation, Success/Failure messages

## SharePoint List Schemas (recommended)

### Tiles / Automation Catalog

| Field | Type |
|-------|------|
| Title | Single line |
| Description | Multiple lines |
| Category | Choice |
| Image | Hyperlink or Image |
| Link | Hyperlink |
| Icon | Single line (Fluent icon name) |
| DisplayOrder | Number |
| IsActive | Yes/No |
| Audience | Choice (multi) |
| Featured | Yes/No |

### Timeline

| Field | Type |
|-------|------|
| Title | Single line |
| Description | Multiple lines |
| StartDate | Date/Time |
| EndDate | Date/Time |
| Status | Choice (Completed, In Progress, Upcoming, On Hold, Cancelled) |
| Category | Choice |
| Owner | Person or Single line |
| Icon | Single line |
| Link | Hyperlink |
| Sequence | Number |
| Audience | Choice (multi) |
| IsActive | Yes/No |

### Accordion / FAQ

| Field | Type |
|-------|------|
| Title | Single line |
| Content | Multiple lines (Rich text) |
| Category | Choice |
| DisplayOrder | Number |
| IsActive | Yes/No |
| Audience | Choice (multi) |
| Link | Hyperlink |
| LastUpdated | Date/Time |

## Security

- Permissions are enforced at the data/API layer (SharePoint list permissions).
- Client-side hiding is never the sole control.
- Power Automate actions validate configuration; secrets are never stored in web-part properties.
- Rich text is sanitized before render.
- URLs for navigation are validated.

## Telemetry

Shared `TelemetryService` tracks:

- WebPartLoaded, ItemViewed, ItemClicked, SearchPerformed, FilterApplied
- LinkOpened, ActionStarted / Completed / Failed
- RequestSubmitted, AutomationLaunched

No sensitive PII is captured.

## Implementation Phases

| Phase | Scope | Status |
|-------|--------|--------|
| 1 — Foundation | Models, List Data Adapter, Config, UI states, Telemetry, Hooks | ✅ Done |
| 2 — Core Renderers | Tiles, Accordion, Timeline (+ Quotes, Spotlight, Q&A) | 🟡 Tiles + Timeline + Accordion renderer |
| 3 — Analytics | Charts, adoption dashboard | Planned |
| 4 — Automation | Automation Hub, Request Center, My Requests, Approval Center | Planned |
| 5 — Enterprise | Full audience, RBAC, localization, a11y validation, performance | Planned |

## Definition of Done (per web part)

- [x] SharePoint List data configurable
- [x] Field mapping configurable
- [x] Filtering / sorting
- [x] Empty / Loading / Error / Unauthorized states
- [x] Responsive layout
- [x] Accessibility foundations (keyboard, ARIA, focus)
- [x] Telemetry
- [ ] Full unit + component test suite
- [ ] Localization packs
- [ ] Security & performance review sign-off

## References (PnP samples used as inspiration)

- [react-tiles-v2](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-tiles-v2)
- [react-quotes](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-quotes)
- [react-chartcontrol](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-chartcontrol)
- [js-employee-spotlight](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/js-employee-spotlight)
- [react-questions-and-answers](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-questions-and-answers)

Reusable logic is refactored into the common architecture rather than copied as isolated implementations.

## License

MIT

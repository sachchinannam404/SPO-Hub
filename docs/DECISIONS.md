# Architecture Decision Log

## ADR-001: Single List Data Adapter

**Decision:** All web parts use `SharePointListDataService` + `useListData` hook.  
**Rationale:** Spec requires no duplicated SharePoint connection, filtering, sorting, or error handling.  
**Consequence:** New renderers only implement mapping + presentation.

## ADR-002: PnPjs over raw REST

**Decision:** Use `@pnp/sp` with SPFx context.  
**Rationale:** Consistent, typed, maintained abstraction; matches enterprise guidance.  
**Consequence:** One dependency; easier expand/select handling.

## ADR-003: Client-side audience filtering (with server top)

**Decision:** Apply audience filter after retrieving a bounded set of items.  
**Rationale:** Multi-choice audience OData is limited; early top + client filter is practical.  
**Consequence:** Do not retrieve unbounded datasets; keep `itemLimit` reasonable.

## ADR-004: No secrets in web part properties

**Decision:** Power Automate integration uses only a configured HTTPS flow URL.  
**Rationale:** Spec security requirements; secrets must not live in the property bag.  
**Consequence:** Flows should use Azure AD or restricted HTTP triggers.

## ADR-005: Fluent UI React 8

**Decision:** Use `@fluentui/react` v8 (compatible with SPFx 1.20).  
**Rationale:** Stable, accessible, already present in SPFx toolchain.  
**Consequence:** Avoid Fluent v9 until SPFx fully supports it in the chosen baseline.

## ADR-006: Timeline is a first-class renderer

**Decision:** Implement Vertical Timeline independently (not derived from Q&A sample).  
**Rationale:** Explicit instruction in the specification.  
**Consequence:** Own data contract, status model, and presentation.

## ADR-007: Incremental delivery

**Decision:** Ship Phase 1 foundation + Tiles + Timeline + Accordion renderer first.  
**Rationale:** Spec implementation order; value early, reduce risk.  
**Consequence:** Remaining renderers (Quotes, Spotlight, Q&A, Charts) follow the same patterns.

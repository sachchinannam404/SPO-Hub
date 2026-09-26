# Admin Forms Integration into SPO-Hub

Source: https://github.com/sachchinannam404/AdminForms

## Location in SPO-Hub

```
src/adminForms/          # Full AdminForms application layer
  models/
  config/requestTypeRegistry.ts
  services/              # RequestService, RoleService, Audit, Notification, …
  components/            # Dashboard, forms, approval, reporting
  hooks/
  context/
  utils/
src/webparts/adminForms/ # SPFx web part entry
scripts/Provision-AdminForms.ps1
docs/ADMIN-FORMS-SETUP.md
```

## Capabilities

| Feature | Component |
|---------|----------|
| Dashboard + KPIs + filters | RequestsDashboard |
| Multi-type dynamic forms | DynamicRequestForm + requestTypeRegistry |
| Stationery / IT / Travel / Leave / Facilities / Procurement | Registry-driven |
| Approvals | ApprovalPanel + RoleService |
| Child line items | ChildItemsList |
| Audit | AuditHistory + AuditService |
| Reporting + CSV | ReportingPanel |
| Power Automate notifications | NotificationService |
| List/group provisioning | SiteProvisioningService / Provision-AdminForms.ps1 |

## Relationship to SPO-Hub list renderers

Admin Forms uses its own `RequestService` + `SharePointRepository` for transactional workflows.

List-driven content web parts (Tiles, Accordion, Quotes, Charts, etc.) continue to use `SharePointListDataService` + `useListData`.

## Property pane

- Power Automate webhook URL
- Ensure lists & groups on load

## Deploy

Same SPO-Hub package (`spo-hub.sppkg`). Add **Admin Forms** from the SPO-Hub group on a page.

Full setup steps: [ADMIN-FORMS-SETUP.md](./ADMIN-FORMS-SETUP.md)

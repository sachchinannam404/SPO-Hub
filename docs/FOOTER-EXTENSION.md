# SPO-Hub Site Footer Extension

SPFx **Application Customizer** that injects an enterprise footer into the modern page **Bottom** placeholder.

## Capabilities

- List-driven or static footer links
- Categories (column layout)
- Copyright line
- Optional colors / inline layout
- Pre-allocated bottom height (120px) to reduce layout shift
- Accessible (`role="contentinfo"`, semantic footer)

## Footer Links list (optional)

| Column | Type |
|--------|------|
| Title | Single line |
| Link or Url | Hyperlink |
| Category | Choice / Text |
| DisplayOrder | Number |
| IsActive | Yes/No |
| OpenInNewTab | Yes/No |

## Deploy

1. `gulp bundle --ship && gulp package-solution --ship`
2. Deploy `spo-hub.sppkg`
3. Activate **SPO-Hub Site Footer Extension** feature (or tenant-wide via ClientSideInstance.xml)
4. Create **Footer Links** list and add items

See `sharepoint/assets/elements.xml` and `ClientSideInstance.xml`.

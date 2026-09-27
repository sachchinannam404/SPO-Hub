# Footer Links — List Schema

SharePoint list used by the **SPO-Hub Site Footer** Application Customizer (`listTitle` default: `Footer Links`).

## List definition

| Property | Value |
|----------|--------|
| **Title** | `Footer Links` |
| **URL** | `Lists/FooterLinks` |
| **Template** | Generic List (`100`) |
| **Versioning** | Optional (off by default) |

## Columns

| Internal name | Display name | Type | Required | Notes |
|---------------|--------------|------|----------|--------|
| **Title** | Title | Single line of text | Yes | Link label |
| **Link** | Link | Hyperlink | Yes | Preferred URL field |
| **Url** | Url | Single line of text | No | Fallback plain URL |
| **Category** | Category | Choice | No | Company, Legal, Support, Resources, Social (+ fill-in) |
| **DisplayOrder** | Display Order | Number | No | Default `100`, sort ascending |
| **IsActive** | Is Active | Yes/No | No | Default Yes; filtered in query |
| **OpenInNewTab** | Open in new tab | Yes/No | No | Default Yes |
| **IconName** | Icon Name | Single line of text | No | Optional Fluent icon |
| **Description** | Description | Multiple lines | No | Tooltip / a11y |

## Views

| View | Filter | Sort |
|------|--------|------|
| **All Links** (default) | — | DisplayOrder, Title |
| **Active Links** | IsActive = Yes | DisplayOrder |

Extension query:

```
filter: (IsActive eq 1 or IsActive eq null)
orderBy: DisplayOrder asc
top: maxLinks (default 12)
select: Id, Title, Link, Url, Category, DisplayOrder, IsActive, OpenInNewTab
```

## Provisioning

```powershell
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/intranet" -Interactive
.\scripts\Provision-FooterLinks.ps1 -AddSampleItems
```

- JSON: `schemas/footer-links-list.schema.json`
- In-browser: `FooterLinksProvisioningService.ensureList(context)`

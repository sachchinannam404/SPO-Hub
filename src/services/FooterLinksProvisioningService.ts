/**
 * Ensures Footer Links list schema exists (site owner / full control).
 */

import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import { Log } from '@microsoft/sp-core-library';

const LOG = 'FooterLinksProvisioningService';

export const FOOTER_LINKS_LIST_TITLE = 'Footer Links';
export const FOOTER_LINKS_LIST_URL = 'Lists/FooterLinks';

const CATEGORY_CHOICES = ['Company', 'Legal', 'Support', 'Resources', 'Social'];

export interface IProvisionResult {
  listCreated: boolean;
  fieldsAdded: string[];
  message: string;
}

export class FooterLinksProvisioningService {
  public static async ensureList(
    context: any,
    listTitle: string = FOOTER_LINKS_LIST_TITLE
  ): Promise<IProvisionResult> {
    const sp: SPFI = spfi().using(SPFx(context));
    const fieldsAdded: string[] = [];
    let listCreated = false;

    try {
      try {
        await sp.web.lists.getByTitle(listTitle)();
      } catch {
        await sp.web.lists.add(listTitle, 'Site footer links for SPO-Hub', 100, false, {
          OnQuickLaunch: false
        });
        listCreated = true;
        Log.info(LOG, `Created list ${listTitle}`);
      }

      const listApi = sp.web.lists.getByTitle(listTitle);

      const ensureField = async (internalName: string, add: () => Promise<unknown>): Promise<void> => {
        try {
          await listApi.fields.getByInternalNameOrTitle(internalName)();
        } catch {
          await add();
          fieldsAdded.push(internalName);
        }
      };

      await ensureField('Link', async () => {
        await listApi.fields.addUrl('Link', { Required: true });
      });
      await ensureField('Url', async () => {
        await listApi.fields.addText('Url', { MaxLength: 2000 });
      });
      await ensureField('Category', async () => {
        await listApi.fields.addChoice('Category', {
          Choices: CATEGORY_CHOICES,
          FillInChoice: true
        });
      });
      await ensureField('DisplayOrder', async () => {
        await listApi.fields.addNumber('DisplayOrder', { DefaultValue: '100' } as any);
      });
      await ensureField('IsActive', async () => {
        await listApi.fields.addBoolean('IsActive', { DefaultValue: '1' } as any);
      });
      await ensureField('OpenInNewTab', async () => {
        await listApi.fields.addBoolean('OpenInNewTab', { DefaultValue: '1' } as any);
      });
      await ensureField('IconName', async () => {
        await listApi.fields.addText('IconName', { MaxLength: 100 });
      });
      await ensureField('Description', async () => {
        await listApi.fields.addMultilineText('Description', { NumberOfLines: 3 });
      });

      try {
        await listApi.views.getByTitle('Active Links')();
      } catch {
        await listApi.views.add('Active Links', false, {
          RowLimit: 100,
          ViewQuery:
            `<Where><Eq><FieldRef Name='IsActive'/><Value Type='Boolean'>1</Value></Eq></Where>` +
            `<OrderBy><FieldRef Name='DisplayOrder' Ascending='TRUE'/></OrderBy>`
        });
        fieldsAdded.push('View:Active Links');
      }

      return {
        listCreated,
        fieldsAdded,
        message: listCreated
          ? `Created "${listTitle}" with schema.`
          : `Ensured schema on "${listTitle}". Fields added: ${fieldsAdded.join(', ') || 'none'}`
      };
    } catch (e: any) {
      Log.error(LOG, e);
      throw e;
    }
  }
}

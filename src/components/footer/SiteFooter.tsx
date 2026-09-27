import * as React from 'react';
import { useMemo } from 'react';
import { Link } from '@fluentui/react/lib/Link';
import { Text } from '@fluentui/react/lib/Text';
import { Stack } from '@fluentui/react/lib/Stack';
import styles from './SiteFooter.module.scss';

export interface IFooterLink {
  id: number;
  title: string;
  url: string;
  category?: string;
  openInNewTab?: boolean;
}

export interface ISiteFooterProps {
  links: IFooterLink[];
  copyrightText: string;
  showCopyright: boolean;
  backgroundColor?: string;
  textColor?: string;
  layout: 'columns' | 'inline';
  siteTitle?: string;
  siteUrl?: string;
}

export const SiteFooter: React.FC<ISiteFooterProps> = ({
  links,
  copyrightText,
  showCopyright,
  backgroundColor,
  textColor,
  layout,
  siteTitle,
  siteUrl
}) => {
  const byCategory = useMemo(() => {
    const map = new Map<string, IFooterLink[]>();
    links.forEach(l => {
      const cat = l.category || 'Links';
      if (!map.has(cat)) map.set(cat, []);
      map.get(cat)!.push(l);
    });
    return map;
  }, [links]);

  const rootStyle: React.CSSProperties = {
    ...(backgroundColor ? { backgroundColor } : {}),
    ...(textColor ? { color: textColor } : {})
  };
  const linkStyle: React.CSSProperties | undefined = textColor ? { color: textColor } : undefined;

  return (
    <footer className={styles.footer} style={rootStyle} role="contentinfo" aria-label="Site footer">
      <div className={styles.inner}>
        {(siteTitle || links.length > 0) && (
          <div className={layout === 'inline' ? styles.inlineRow : styles.columns}>
            {siteTitle && (
              <div className={styles.brand}>
                {siteUrl ? (
                  <Link href={siteUrl} className={styles.brandLink} style={linkStyle}>{siteTitle}</Link>
                ) : (
                  <Text className={styles.brandText}>{siteTitle}</Text>
                )}
              </div>
            )}

            {layout === 'columns' && byCategory.size > 0 && (
              <div className={styles.columnGrid}>
                {Array.from(byCategory.entries()).map(([category, catLinks]) => (
                  <div key={category} className={styles.column}>
                    <Text className={styles.columnTitle} as="h3">{category}</Text>
                    <ul className={styles.linkList}>
                      {catLinks.map(link => (
                        <li key={link.id}>
                          <Link
                            href={link.url}
                            target={link.openInNewTab ? '_blank' : undefined}
                            rel={link.openInNewTab ? 'noopener noreferrer' : undefined}
                            className={styles.link}
                            style={linkStyle}
                          >
                            {link.title}
                          </Link>
                        </li>
                      ))}
                    </ul>
                  </div>
                ))}
              </div>
            )}

            {layout === 'inline' && links.length > 0 && (
              <nav className={styles.inlineNav} aria-label="Footer links">
                {links.map((link, idx) => (
                  <React.Fragment key={link.id}>
                    {idx > 0 && <span className={styles.sep} aria-hidden="true">|</span>}
                    <Link
                      href={link.url}
                      target={link.openInNewTab ? '_blank' : undefined}
                      rel={link.openInNewTab ? 'noopener noreferrer' : undefined}
                      className={styles.link}
                      style={linkStyle}
                    >
                      {link.title}
                    </Link>
                  </React.Fragment>
                ))}
              </nav>
            )}
          </div>
        )}

        {showCopyright && (
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap className={styles.bottomBar} tokens={{ childrenGap: 8 }}>
            <Text className={styles.copyright} variant="small">{copyrightText}</Text>
            <Text className={styles.powered} variant="small">Powered by SPO-Hub</Text>
          </Stack>
        )}
      </div>
    </footer>
  );
};

export default SiteFooter;

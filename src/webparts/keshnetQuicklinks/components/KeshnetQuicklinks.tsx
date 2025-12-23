import * as React from 'react';
import styles from './KeshnetQuicklinks.module.scss'; // Ensure this path is correct based on your project structure
import type { IKeshnetQuicklinksProps } from './IKeshnetQuicklinksProps';
import {
  getLatestQuickLink,
  getCurrentUser,
  getSPListItemsById,
} from "../Utility/utils"

interface IKeshnetQuicklinksState {
  listData: any[];
  currentUser: any;
  employeeList: any[];
  loading: boolean;
}

export default class KeshetnetQuickLinks extends React.Component<IKeshnetQuicklinksProps, IKeshnetQuicklinksState, {}> {

  constructor(props: IKeshnetQuicklinksProps) {
    super(props);
    this.state = {
      listData: [],
      currentUser: null,
      employeeList: [],
      loading: true,
    };
  }

  async componentDidMount() {
    try {
      const { selectedList, PhoneBookListForQuickLink, context } = this.props;

      const [listData, currentUser, employeeList] = await Promise.all([
        selectedList
          ? getLatestQuickLink({ title: selectedList })
          : Promise.resolve([]),
        getCurrentUser(),
        PhoneBookListForQuickLink
          ? getSPListItemsById(PhoneBookListForQuickLink, context)
          : Promise.resolve([]),
      ]);

      this.setState({
        listData,
        currentUser,
        employeeList,
        loading: false,
      });
    } catch (error) {
      console.error("Error in data fetch:", error);
      this.setState({ loading: false });
    }
  }

  // Click handler
  private handleLinkClick = (e: React.MouseEvent<HTMLAnchorElement>, url: string, openInNewTab: boolean) => {
    e.preventDefault();
    if (!url) return;

    if (openInNewTab) {
      window.open(url, '_blank', 'noopener,noreferrer');
    } else {
      window.location.href = url;
    }
  };

  public render(): React.ReactElement<IKeshnetQuicklinksProps> {
    const { listData, currentUser, employeeList } = this.state;

    // Create a sorted copy for rendering
    const displayData = [...listData]?.sort((a, b) => {
      const orderA = a.order0 !== null && a.order0 !== undefined && a.order0 !== '' ? Number(a.order0) : Number.MAX_SAFE_INTEGER;
      const orderB = b.order0 !== null && b.order0 !== undefined && b.order0 !== '' ? Number(b.order0) : Number.MAX_SAFE_INTEGER;
      return orderA - orderB;
    });

    const firstName = employeeList?.filter((x: { Email: any; }) => { return x.Email?.toLowerCase().trim() == currentUser.email?.toLowerCase().trim() })[0]?.spHebFirstname

    return (
      <section className={styles.quickLinks_wrap} dir="rtl" lang="he">
        <main className={styles.quicklinks_content}>
          <div className={styles.quicklinks_MainDivHeightWidth}>
            {!this.state.loading && (
              <div className={styles.QuickLinksWebPart_greeting}>
                היי<span className={styles['QuickLinksWebPart_bold-part']}>{firstName || ''}</span>
              </div>
            )}


            <div className={styles.quickLinks_actions}>
              {displayData.map((item: any, index: any) => {
                const { Title, URL, OpenInNewTab, AttachmentFiles } = item;
                const iconFile = AttachmentFiles?.[0]?.ServerRelativeUrl
                return (
                  <a
                    key={index}
                    className={`${styles.quickLinks_action} ${styles['icon-a']}`}
                    title={Title}
                    target={OpenInNewTab ? '_blank' : '_self'}
                    onClick={(e) => this.handleLinkClick(e, URL, OpenInNewTab)}
                  >
                    <span className={styles.quickLinks_icon} aria-hidden="true">
                      {iconFile && <img style={{ height: '54px', width: '54px' }} src={iconFile} alt={Title} />}
                    </span>
                    <span className={styles.quickLinks_label}>{Title}</span>
                  </a>
                );
              })}
            </div>

          </div>
        </main>
      </section>
    );
  }
}

import * as React from 'react';
import styles from './KeshnetBirthday.module.scss';
import type { IKeshnetBirthdayProps } from './IKeshnetBirthdayProps';
import { getAllImages, getListItemsFormainList } from '../Utility/utils';

interface IGalleryCard {
  FileRef: any;
  FileLeafRef?: string;
  EmployeeName?: { EMail: string };
}

const KeshnetBirthday: React.FC<IKeshnetBirthdayProps> = (props) => {
  const [Images, setImages] = React.useState<IGalleryCard[]>([]);
  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [listData, setListData] = React.useState<any[]>([]);
  const [isListLoaded, setIsListLoaded] = React.useState(false); // new loading flag

  const { selectedList, ImageLibrary, Duration } = props;

  const fetchData = async (): Promise<void> => {
    try {
      const images: any = await getAllImages(ImageLibrary);
      setImages(images);
    } catch (error) {
      console.error("Error fetching gallery items:", error);
    }
  };

  const fetchListData = async () => {
    if (!selectedList) return;
    setIsListLoaded(false); // start loading
    try {
      const data = await getListItemsFormainList({ title: selectedList });
      setListData(data || []);
    } catch (error) {
      console.error("Error fetching list data:", error);
    } finally {
      setIsListLoaded(true); // finished loading (success or fail)
    }
  };

  React.useEffect(() => {
    fetchData();
    fetchListData();
  }, [selectedList, ImageLibrary]);

  React.useEffect(() => {
    if (!listData || listData.length === 0) return;

    const interval = setInterval(() => {
      setCurrentIndex((prev) => (prev + 1) % listData.length);
    }, (Number(Duration) || 5) * 1000);

    return () => clearInterval(interval);
  }, [listData, Duration]);

  const currentItem =
    listData && listData.length > 0 ? listData[currentIndex] : null;

  const backgroundImageUrl = currentItem
    ? require('../assets/background.png')
    : require('../assets/Frame1702.png');

  const matchedImage = Images.find(
    (i) =>
      i.EmployeeName &&
      i.EmployeeName?.EMail?.toLowerCase().trim() ===
      currentItem?.Email?.toLowerCase().trim()
  );

  const attachmentUrl =
    matchedImage && currentItem
      ? matchedImage.FileRef
      : require('../assets/DefaultIMG.png');

  const displayName = currentItem ? currentItem[props.fullname] : "";
  const department = currentItem ? currentItem[props.department] : "";
  const company = currentItem ? currentItem[props.company] : "";

  return (
    <div className={styles.birthdayWebPart_wrap} dir="rtl" lang="he">
      <div className={styles.birthdayWebPart_content}>

        {/* Hide title until list data is fully loaded */}
        <div className={styles.birthdayWebPart_center_title}>
          {isListLoaded && (
            <>
              <span className={styles.birthdayWebPart_bold_text}>היום</span>
              <span className={styles.birthdayWebPart_accent}>יומולדת</span>
            </>
          )}
        </div>


        <div className={styles.birthdayWebPart_birthday_block}>
          <div className={styles.birthdayWebPart_right_column}>
            <div className={styles.birthdayWebPart_photo_wrap}>
              <div className={styles.birthdayWebPart_photo}>
                <img
                  style={{
                    borderTopRightRadius: '8px',
                    borderBottomRightRadius: '8px',
                  }}
                  src={attachmentUrl}
                  alt={displayName}
                />
              </div>
            </div>
            <div
              className={styles.birthdayWebPart_card_left}
              style={{ backgroundImage: `url(${backgroundImageUrl})` }}
            >
              <div className={styles.birthdayWebPart_card_content}>
                <div className={styles.birthdayWebPart_name}>{displayName}</div>
                <div className={styles.birthday_Department_Copmany}>
                  <div className={styles.birthdayWebPart_department}>
                    {department}
                  </div>
                  <div className={styles.birthdayWebPart_company}>
                    {company}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default KeshnetBirthday;

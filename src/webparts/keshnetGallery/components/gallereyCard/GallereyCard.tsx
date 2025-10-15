/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import * as moment from 'moment';
import styles from './GalleryCard.module.scss';
import staticImage from '../../assets/default.svg';
import gradImage from '../../assets/grad.svg';


interface IGalleryCard {
    item: any;
    OpenCurrntCard: any;
    Images: any[];
    Template: any[];
    Employees: any[];
}

const GalleryCard: React.FC<IGalleryCard> = ({ item, OpenCurrntCard, Images, Template, Employees }) => {
  const matchedImage = Images.find((i: any) =>
    (i?.EmployeeName?.EMail ? i.EmployeeName.EMail.toLowerCase().trim() : "") ===
    (item?.Employee?.EMail ? item.Employee.EMail.toLowerCase().trim() : "")
  );

  const imageUrl = matchedImage ? matchedImage.FileRef : staticImage;
  const formattedDate = item?.Date ? moment(item.Date).format("DD.MM.YY") : "";

  const currentEmploye = Employees.find((a: any) =>
    (a?.Email ? a.Email.toLowerCase().trim() : "") ===
    (item?.Employee?.EMail ? item.Employee.EMail.toLowerCase().trim() : "")
  );

  const TemplateItem = Template.find((a: any) =>
    (a?.Category?.trim() ?? "") === (item?.Category?.trim() ?? "")
  );

  const cardImage =
    TemplateItem?.CardType === "Group"
      ? (TemplateItem?.ImagePath || staticImage)
      : imageUrl;

  return (
    <div className={styles.customCard} onClick={() => OpenCurrntCard(true, item)}>
      <div className={styles.iconLabel}>
        {TemplateItem?.IconPath ? (
          <img src={TemplateItem.IconPath} alt="icon" />
        ) : (
          <div></div>
        )}
        <p>{item?.Category}</p>
      </div>

      <div className={styles.cardCustomBorder}>
        <div className={styles.color5} />
        <div className={styles.color7} />
        <div className={styles.color6} />
        <div className={styles.color8} />
        <div className={styles.color2} />
        <div className={styles.color3} />
        <div className={styles.color1} />
        <div className={styles.color4} />
      </div>

      <div className={styles.imageContainer}>
        <img src={cardImage} alt="Card Image" className={styles.backGalleryImage} />
        <img src={gradImage} alt="grad" className={styles.imageSpanGradient} />
      </div>

      <div className={styles.overlay}>
        <span className={styles.overlaySpanGradient} />
        <div className={styles.bottomText}>
          <span className={styles.number}>{formattedDate}</span>
          <span className={styles.label}>
            {TemplateItem?.CardType === "Personal"
              ? currentEmploye?.spFullName
              : (item?.Title || item?.Employee?.Title)}
          </span>
        </div>
      </div>
    </div>
  );
};


export default GalleryCard;
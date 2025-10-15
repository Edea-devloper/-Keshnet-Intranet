import * as React from 'react';
import styles from './CardModal.module.scss';
import staticImage from '../../assets/default.svg';

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type

export interface IUserModalProps {
    isOpen: boolean;
    onClose: () => void;
    item: any;
    Images: any[];
    TemplateItems: any[];
    Employees: any[];
}

const CardModal: React.FC<IUserModalProps> = ({
    isOpen,
    onClose,
    item,
    Images,
    TemplateItems,
    Employees
}) => {

    if (!isOpen) return null;

    // Safe compare with ?? "" before toLowerCase/trim
    const matchedImage = Images.find((i: any) =>
        (i?.EmployeeName?.EMail ?? "").toLowerCase().trim() === (item?.Employee?.EMail ?? "").toLowerCase().trim()
    );

    const imageUrl = matchedImage ? matchedImage.FileRef : staticImage;

    const currentEmploye = Employees.find((a: any) =>
        (a?.Email ?? "").toLowerCase().trim() === (item?.Employee?.EMail ?? "").toLowerCase().trim()
    );

    const TemplateItem = TemplateItems.find((a: any) =>
        (a?.Category ?? "").trim() === (item?.Category ?? "").trim()
    );

    const CurrentEmployee = Employees.find((a: any) =>
        (a?.Email ?? "").trim() === (item?.Employee?.EMail ?? "").trim()
    );

    return (
        <div className={styles.modalOverlay} onClick={onClose}>
            <div className={styles.modalContainer} onClick={(e) => e.stopPropagation()}>
                <button className={styles.closeButton} onClick={onClose}>
                    <span aria-hidden="true">&times;</span>
                </button>

                <div className={styles.content}>
                    <div className={styles.leftContent}>
                        <div className={styles.contentContainer}>
                            <div className={styles.title}>
                                {TemplateItem && TemplateItem.IconPath ? (
                                    <img
                                        className={styles.galleryicon}
                                        src={TemplateItem.IconPath}
                                        alt="icon"
                                    />
                                ) : (
                                    <div></div>
                                )}
                                {item?.Category}
                            </div>

                            <div className={styles.cardContent}>
                                <div className={styles.cardTitleDetails}>
                                    <div className={styles.name}>
                                        {TemplateItem && TemplateItem.CardType === "Personal"
                                            ? currentEmploye?.spFullName
                                            : (item?.Title ? item?.Title : item?.Employee?.Title)}
                                    </div>
                                    {TemplateItem && TemplateItem.CardType === "Personal" ? (
                                        <div
                                            className={styles.department}
                                            title={CurrentEmployee?.Title ? CurrentEmployee?.Title : item?.Employee?.JobTitle}
                                        >
                                            {CurrentEmployee?.Title ? CurrentEmployee?.Title : item?.Employee?.JobTitle}
                                        </div>
                                    ) : null}
                                </div>
                                <div className={styles.message}>{item?.Description}</div>
                            </div>
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
                    </div>

                    <div className={styles.rightContent}>
                        <img
                            src={(TemplateItem && TemplateItem.CardType === "Group" ? TemplateItem.ImagePath : imageUrl)}
                            alt="Card Image"
                        />
                    </div>
                </div>
            </div>
        </div>
    );
};

export default CardModal;

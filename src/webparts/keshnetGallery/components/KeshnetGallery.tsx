import * as React from 'react';
// import styles from './KeshnetGallery.module.scss';
import styles from './gallereyCard/GalleryCard.module.scss';
import type { IKeshnetGalleryProps } from './IKeshnetGalleryProps';
import GallereyCard from './gallereyCard/GallereyCard';
import { getAllImages, getEmployees, getSPListItemsById, GetTemplateList } from '../services/SPService';
import CardModal from './modal/CardModal';
import NavigationLeft from '../assets/navigation_Left.svg';
import NavigationRight from '../assets/navigation_Right.svg';

const KeshnetGallery: React.FC<IKeshnetGalleryProps> = ({ _GalleryLists, context, ImageLibrary, TemplateList, EmployeeList, AssetFolderPath }) => {

  const [galleryItems, setGalleryItems] = React.useState([]); // Card data from SP list
  const [Images, SetImages] = React.useState([]);
  const [TemplateItems, settemplateItems] = React.useState([] as any);
  const [Employees, setEmployees] = React.useState([] as any);
  const [isOpen, setIsOpen] = React.useState(false); // Open/close modal
  const [currntItem, setCurrntItem] = React.useState({}); // Current card item for modal
  const [isDataLoaded, setIsDataLoaded] = React.useState(false);

  // State to track scroll availability (RTL adjusted)
  const [canScrollPrev, setCanScrollPrev] = React.useState(false); // visually left
  const [canScrollNext, setCanScrollNext] = React.useState(false); // visually right

  React.useEffect(() => {
    const style = `
      /* Regular (400) */
      @font-face {
        font-family: 'Keshet-12';
        src: url("${AssetFolderPath}Keshet-12-Regular.woff2") format('woff2');
        font-weight: 400;
        font-style: normal;
        font-display: swap;
      }

      /* Light (300) */
      @font-face {
        font-family: 'Keshet-12';
        src: url("${AssetFolderPath}Keshet-12-Light.woff2") format('woff2');
        font-weight: 300;
        font-style: normal;
        font-display: swap;
      }

      /* ExtraLight (200) */
      @font-face {
        font-family: 'Keshet-12';
        src: url("${AssetFolderPath}Keshet-12-ExtraLight.woff2") format('woff2');
        font-weight: 200;
        font-style: normal;
        font-display: swap;
      }

      /* SemiBold (600) */
      @font-face {
        font-family: 'Keshet-12';
        src: url("${AssetFolderPath}Keshet-12-SemiBold.woff2") format('woff2');
        font-weight: 600;
        font-style: normal;
        font-display: swap;
      }

      /* Bold (700) */
      @font-face {
        font-family: 'Keshet-12';
        src: url("${AssetFolderPath}Keshet-12-Bold.woff2") format('woff2');
        font-weight: 700;
        font-style: normal;
        font-display: swap;
      }

      /* ExtraBold (800) */
      @font-face {
        font-family: 'Keshet-12';
        src: url("${AssetFolderPath}Keshet-12-ExtraBold.woff2") format('woff2');
        font-weight: 800;
        font-style: normal;
        font-display: swap;
      }

      /* Apply globally */
      body, p, div, span, a, h1, h2, h3 {
        font-family: 'Keshet-12' !important;
      }
    `;

    // Create and inject <style>
    const styleElement = document.createElement("style");
    styleElement.id = "keshet-fonts";
    styleElement.innerText = style;
    document.head.appendChild(styleElement);

  }, []);


  React.useEffect(() => {
    const loadData = async () => {
      try {
        const [img, data, emp, items]: [any, any, any, any] = await Promise.all([
          getAllImages(ImageLibrary),
          GetTemplateList(TemplateList),
          getEmployees(EmployeeList),
          getSPListItemsById(_GalleryLists),
        ]);

        const sortedItems = items?.sort((a: any, b: any) => {
          const orderA = a.order0 !== null && a.order0 !== undefined && a.order0 !== '' ? Number(a.order0) : Number.MAX_SAFE_INTEGER;
          const orderB = b.order0 !== null && b.order0 !== undefined && b.order0 !== '' ? Number(b.order0) : Number.MAX_SAFE_INTEGER;

          return orderA - orderB;
        });


        SetImages(img);
        settemplateItems(data);
        setEmployees(emp);
        setGalleryItems(sortedItems);

        setIsDataLoaded(true);
      } catch (err) {
        console.error('Error loading gallery data', err);
      } finally {
        setTimeout(checkScroll, 100);
      }
    };

    loadData();
  }, []);

  const checkScroll = () => {
    const container = document.getElementById('carousel');
    if (container) {
      // In RTL, scrollLeft is negative in most browsers (Chrome/Edge)
      const maxScroll = container.scrollWidth - container.clientWidth;
      const scrollPos = Math.abs(container.scrollLeft);

      setCanScrollPrev(scrollPos > 0); // Show previous if not at start
      setCanScrollNext(scrollPos < maxScroll - 1); // Show next if not at end
    }
  };

  // Attach scroll listener
  React.useEffect(() => {
    const container = document.getElementById('carousel');
    if (container) {
      checkScroll();
      container.addEventListener('scroll', checkScroll);
      return () => container.removeEventListener('scroll', checkScroll);
    }
  }, [galleryItems]);


  const OpenAndCloseCurrntCard = (isOpen: boolean, CurrntCardItem: any) => {
    setIsOpen(isOpen);
    setCurrntItem(CurrntCardItem);
  };

  const scrollCarousel = (direction: number) => {
    const container = document.getElementById('carousel');
    if (container) {
      const scrollAmount = 228; // card width + gap

      // In RTL, direction is reversed: next = -scrollAmount, prev = +scrollAmount
      container.scrollBy({
        left: -direction * scrollAmount,
        behavior: 'smooth'
      });

      setTimeout(checkScroll, 300); // Re-check after scroll animation
    }
  };

  return (
    <div dir="rtl" style={{ width: '99.4%' }}>
      <div style={{ display: 'flex', justifyContent: 'center', width: '100%' }}>
        <div style={{ width: '960px' }} className={styles.customGalleryContainer}>
          <div className={styles.main}>
            {isDataLoaded && (
              <div className={`${styles.carouselWrapperContainer} ${styles.sectionTitle}`}>
                מה קורה <span>בקשת</span>
              </div>
            )}


            <div className={`${styles.carouselWrapperContainer} ${styles.carouselWrapper}`}>

              {/* Next button (visually right in RTL) */}
              {canScrollNext && (
                <img src={NavigationLeft} alt="NavigationLeft" className={`${styles.carouselButton} ${styles.left}`} onClick={() => scrollCarousel(1)} />
              )}

              <div className={styles.cardCarousel} id="carousel">
                <div className={styles.cardGrid}>
                  {galleryItems.length > 0 ? (
                    galleryItems.map((g: any) => (
                      <GallereyCard key={g.Id} item={g} OpenCurrntCard={OpenAndCloseCurrntCard} Images={Images} Template={TemplateItems} Employees={Employees} />
                    ))
                  ) : (
                    <p className={styles.noData}></p>
                  )}
                </div>
              </div>

              {/* Previous button (visually left in RTL) */}
              {canScrollPrev && (
                <img src={NavigationRight} alt="NavigationRight" className={`${styles.carouselButton} ${styles.right}`} onClick={() => scrollCarousel(-1)} />
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Card Modal */}
      <CardModal
        isOpen={isOpen}
        onClose={() => OpenAndCloseCurrntCard(false, {})}
        item={currntItem}
        Images={Images}
        TemplateItems={TemplateItems}
        Employees={Employees}
      />
    </div>
  );
};

export default KeshnetGallery;

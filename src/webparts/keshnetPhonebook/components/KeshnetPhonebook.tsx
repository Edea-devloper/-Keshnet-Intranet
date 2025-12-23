import * as React from 'react';
import { useState } from 'react';
import styles from './KeshnetPhonebook.module.scss';
import { IKeshnetPhonebookProps } from './IKeshnetPhonebookProps';
import { getSPListItemsById } from '../Utility/utils';
import userImg from '../assets/user.svg';
import layerImg from '../assets/Layer_1.png';
import Less from '../assets/Less.png';
import Greater from '../assets/Greater.png';
import { getAllImages } from '../Utility/utils';


interface Employee {
  id: number;
  Email: string;
  Phone: string;
  extension: string;
  Department: string;
  Position: string;
  Location: string;
  building: string;
  jobTitle: string;
  firstName: string;
  lastName: string;
  avatar: string;
  frame: string;
  layer?: string; // Optional: only used in modal
  [key: string]: string | number | boolean | Date | undefined; // Allow string-based property access
}

interface Column {
  key: string;
  label: string;
  style?: React.CSSProperties;
}
interface SPImage {
  FileLeafRef: string;
  FileRef: string;
  EmployeeName?: { Email: string }
}


// const layer = require('../assets/layer.png'); // Layer image for modal

const KeshetNet: React.FC<IKeshnetPhonebookProps> = (props) => {
  const [currentPage, setCurrentPage] = useState(1);
  const [searchTerm, setSearchTerm] = useState('');
  const itemsPerPage = Number(props.maxRows) || 10; // ensures it's always a number
  const [activeTab, setActiveTab] = React.useState('עובדי קשת'); // Default selected tab
  const [selectedEmployee, setSelectedEmployee] = useState<Employee | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [phoneBookData, setPhoneBookData] = useState<any[]>([]);
  const [Images, SetImages] = useState<SPImage[]>([]);

  React.useEffect(() => {
    const style = `
        /* Regular (400) */
        @font-face {
          font-family: 'Keshet-12';
          src: url("${props.AssetFolderPath}Keshet-12-Regular.woff2") format('woff2');
          font-weight: 400;
          font-style: normal;
          font-display: swap;
        }
  
        /* Light (300) */
        @font-face {
          font-family: 'Keshet-12';
          src: url("${props.AssetFolderPath}Keshet-12-Light.woff2") format('woff2');
          font-weight: 300;
          font-style: normal;
          font-display: swap;
        }
  
        /* ExtraLight (200) */
        @font-face {
          font-family: 'Keshet-12';
          src: url("${props.AssetFolderPath}Keshet-12-ExtraLight.woff2") format('woff2');
          font-weight: 200;
          font-style: normal;
          font-display: swap;
        }
  
        /* SemiBold (600) */
        @font-face {
          font-family: 'Keshet-12';
          src: url("${props.AssetFolderPath}Keshet-12-SemiBold.woff2") format('woff2');
          font-weight: 600;
          font-style: normal;
          font-display: swap;
        }
  
        /* Bold (700) */
        @font-face {
          font-family: 'Keshet-12';
          src: url("${props.AssetFolderPath}Keshet-12-Bold.woff2") format('woff2');
          font-weight: 700;
          font-style: normal;
          font-display: swap;
        }
  
        /* ExtraBold (800) */
        @font-face {
          font-family: 'Keshet-12';
          src: url("${props.AssetFolderPath}Keshet-12-ExtraBold.woff2") format('woff2');
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
    if (!props.themeFont) return;

    const root = document.documentElement;
    root.style.setProperty('--themeFontFamily', props.themeFont || 'Segoe UI');
  }, [props.themeFont]);

  // Reset page when tab, search, or rows per page change
  React.useEffect(() => {
    setCurrentPage(1);
  }, [activeTab, searchTerm, itemsPerPage]);

  // Reset searchTerm ONLY when tab changes
  React.useEffect(() => {
    setSearchTerm('');
  }, [activeTab]);

  const isValidDateValue = (value: any): boolean => {
    if (value instanceof Date) return true;
    if (typeof value !== "string") return false;

    const str = value.trim();

    // Supported formats:
    const simplePattern = /^\d{4}-\d{2}-\d{2}$/; // 2025-08-03
    const commonPattern = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/; // 03/08/2025
    const isoPattern = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z?$/; // 2025-08-03T07:00:00Z

    if (!simplePattern.test(str) && !commonPattern.test(str) && !isoPattern.test(str)) {
      return false;
    }

    const parsed = Date.parse(value);
    return !isNaN(parsed);
  };


  const formatDate = (dateValue?: string | Date): string => {
    if (!dateValue) return '';
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return ''; // invalid date safety
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  const dynamicColumns: Column[] = React.useMemo(() => {
    let selectedCols: string[] = [];

    // Choose columns based on active tab
    if (activeTab === "עובדי קשת") {
      selectedCols = [...(props.selectedColumnsFirst || [])];
    } else if (activeTab === "עובדי הפקות") {
      selectedCols = [...(props.selectedColumnsSecond || [])];
    } else if (activeTab === "KI") {
      selectedCols = [...(props.selectedColumnsThird || [])];
    }

    return selectedCols.map((colKey) => {
      const columnInfo = props.listColumns?.find(c => c.key === colKey);
      return {
        key: colKey,
        label: columnInfo ? columnInfo.text : colKey,
        style:
          colKey === "spFullName" || colKey === "SPMobilePhone"
            ? { fontWeight: "bold" }
            : colKey === "JobTitle"
              ? { width: "200px", fontWeight: "normal" }
              : colKey === "Department"
                ? { width: "120px", fontWeight: "normal" }
                : { fontWeight: "normal" },
      };
    });
  }, [
    activeTab,
    props.selectedColumnsFirst,
    props.selectedColumnsSecond,
    props.selectedColumnsThird,
    props.listColumns
  ]);



  const handleRowClick = (employee: Employee) => {
    setSelectedEmployee(employee);
    setIsModalOpen(true);
  };

  React.useEffect(() => {
    const fetchPhoneBookData = async () => {
      try {
        const images: any = await getAllImages(props.ImageLibrary, props.context);
        SetImages(images);

        const items = await getSPListItemsById(props.PhoneBookList, props.context);

        // format Birthdate before saving to state
        let formattedItems = items.map((item: any) => ({
          ...item,
          BirthDate: formatDate(item.BirthDate)
        }));

        // Sort alphabetically by selected column (from property pane)
        if (props.selectordercolumn) {
          formattedItems = [...formattedItems].sort((a, b) => {
            const valA = normalizeForSort(a[props.selectordercolumn]);
            const valB = normalizeForSort(b[props.selectordercolumn]);
            return valA.localeCompare(valB, 'he', { sensitivity: 'base' });
          });
        }

        setPhoneBookData(formattedItems);
      } catch (error) {
        console.error('Error fetching data:', error);
      }
    };

    void fetchPhoneBookData();
  }, [props.PhoneBookList, props.context]);


  const normalizeForSort = (text: any): string => {
    if (!text) return '';
    return String(text)
      .normalize("NFKC")                 // normalize Unicode
      .replace(/[\u0591-\u05C7]/g, "")   // remove Hebrew nikud
      .replace(/\u200F|\u200E/g, "")     // remove RTL/LTR invisible marks
      .trim();                           // remove spaces
  };

  const filteredEmployees = phoneBookData
    .filter(employee => {
      if (activeTab === "עובדי קשת") {
        return (
          employee.sharepoint_status === "active" &&
          [
            "SAP",
            "Freelancer-Promo&Post",
            "Freelancer-FreeTv",
            "Freelancer-KeshetDigital",
            "Freelancer-Keshet"
          ].includes(employee.SPbelong)
        );
      }

      if (activeTab === "עובדי הפקות") {
        return (
          employee.sharepoint_status === "active" &&
          employee.SPbelong === "הפקות"
        );
      }

      if (activeTab === "KI") {
        return (
          employee.sharepoint_status === "active" &&
          [
            "קשת אינטרנשיונל",
            "Keshet International"
          ].includes(employee.spCompanyName)
        );
      }

      return employee.Category === activeTab;
    })
    .filter(employee => {
      if (!searchTerm.trim()) return true; // no filter if empty

      // --- Improved normalizeText handles both RTL/LTR safely ---
      const normalizeText = (text: string): string => {
        if (!text) return '';
        return text
          .toString()
          .normalize("NFD") // split combined unicode (accents)
          .replace(/[\u0300-\u036f]/g, "") // remove accents
          .replace(/[\u200E\u200F]/g, "") // remove RTL/LTR invisible marks
          .replace(/[’‘']/g, "'") // normalize quotes
          .replace(/[״”“]/g, '"') // normalize double quotes
          .replace(/\s+/g, " ") // collapse multiple spaces
          .trim()
          .toLowerCase();
      };

      // Normalize the search term itself
      const search = normalizeText(searchTerm);

      // Determine which set of columns to search based on active tab
      const activeColumns =
        activeTab === "עובדי קשת"
          ? props.selectedColumnsFirst || []
          : activeTab === "עובדי הפקות"
            ? props.selectedColumnsSecond || []
            : activeTab === "KI"
              ? props.selectedColumnsThird || []
              : [];

      // Match search text against the active tab's selected columns
      return activeColumns.some(colKey => {
        const val = employee[colKey];
        if (typeof val !== "string") return false;

        // Normalize the value for consistent comparison
        const normalizedValue = normalizeText(val);

        return normalizedValue.includes(search);
      });
    });


  const totalPages = Math.ceil(filteredEmployees.length / itemsPerPage);
  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(1);
    }
  }, [totalPages]);

  // Resort data whenever property pane "selectordercolumn" changes
  React.useEffect(() => {
    if (!props.selectordercolumn) return;

    setPhoneBookData(prevData =>
      [...prevData].sort((a, b) => {
        const valA = normalizeForSort(a[props.selectordercolumn]);
        const valB = normalizeForSort(b[props.selectordercolumn]);
        return valA.localeCompare(valB, 'he', { sensitivity: 'base' });
      })
    );

  }, [props.selectordercolumn]);


  const startIndex = (currentPage - 1) * itemsPerPage;
  const displayedEmployees = filteredEmployees.slice(startIndex, startIndex + itemsPerPage);

  const handlePageChange = (page: number) => {
    setCurrentPage(page);
  };

  const renderPagination = () => {
    const pages = [];
    const maxVisiblePages = 12;
    const startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
    const endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

    // Previous button
    pages.push(
      <button
        key="prev"
        className={`${styles.pageButton} ${styles.arrowButton} ${currentPage === 1 ? styles.disabled : ''}`}
        onClick={() => currentPage > 1 && handlePageChange(currentPage - 1)}
        disabled={currentPage === 1}
      >
        <img src={Less} alt="Previous" className={styles.icon} />
      </button>
    );

    // Page numbers...
    for (let i = endPage; i >= startPage; i--) {
      pages.push(
        <button
          key={i}
          className={`${styles.pageButton} ${i === currentPage ? styles.active : ''}`}
          onClick={() => handlePageChange(i)}
        >
          <div className={styles.paginationNumber}>{i}</div>
        </button>
      );
    }

    // Next button
    pages.push(
      <button
        key="next"
        className={`${styles.pageButton} ${styles.arrowButton} ${currentPage === totalPages ? styles.disabled : ''}`}
        onClick={() => currentPage < totalPages && handlePageChange(currentPage + 1)}
        disabled={currentPage === totalPages}
      >
        <img src={Greater} alt="Next" className={styles.icon} />
      </button>
    );


    return pages;
  };

  const getImageUrl = (employee: any) => {
    const empEmail = employee?.Email ? employee.Email.toLowerCase().trim() : "";
    const matchedImage = Images?.find((i: any) => {
      const imgEmail = i?.EmployeeName?.EMail ? i.EmployeeName.EMail.toLowerCase().trim() : "";
      return imgEmail && empEmail && imgEmail === empEmail;
    });
    return matchedImage?.FileRef || userImg;
  };


  return (
    <div className={styles.keshetNetContainer}>
      <div className={styles.keshetNet}>
        <div className={styles.tabs}>
          <div
            className={`${styles.tab} ${activeTab === 'עובדי קשת' ? styles.active : styles.inactive}`}
            onClick={() => { setActiveTab('עובדי קשת'); setSearchTerm(''); }}
          >
            עובדי קשת
          </div>
          <div
            className={`${styles.tab} ${activeTab === 'עובדי הפקות' ? styles.active : styles.inactive}`}
            onClick={() => { setActiveTab('עובדי הפקות'); setSearchTerm(''); }}
          >
            עובדי הפקות
          </div>
          <div
            className={`${styles.tab} ${activeTab === 'KI' ? styles.active : styles.inactive}`}
            onClick={() => { setActiveTab('KI'); setSearchTerm(''); }}
          >
            KI
          </div>

        </div>

        <div className={styles.header}>
          <div className={styles.searchContainer}>
            <input
              type="text"
              placeholder="חיפוש שם"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className={styles.searchInput}
            />
          </div>
        </div>

        <div className={styles.tableContainer}>
          <table className={styles.employeeTable} cellPadding={0} cellSpacing={0}>
            <thead>
              <tr>
                <th></th> {/* Avatar column */}
                {dynamicColumns.map((col) => {
                  let label = col.label;

                  // Change labels only when activeTab === "עובדי הפקות"
                  if (activeTab === "עובדי הפקות") {
                    if (col.key === "spCompanyName") label = "יחידה"; // Unit
                    if (col.key === "SPmanager") label = "הפקות הבית"; // Department
                    if (col.key === "SPsubDepartment") label = "תת מחלקה"; // Sub-department
                  }

                  return (
                    <th key={col.key} style={col.style}>
                      {label}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {displayedEmployees.map((employee) => {
                return (
                  <tr
                    key={employee.id}
                    onClick={() => handleRowClick(employee)}
                    className={styles.clickableRow}
                  >
                    <td style={{ width: '50px', border: 'none' }} >
                
                      <img
                        style={{ height: '60px', width: '54px', marginBottom: '-6px', marginTop: '-1px', marginLeft: '-1px', marginRight: '-1px', objectFit: 'fill' }}
                        src={getImageUrl(employee) || userImg}
                        alt={`${employee.firstName} ${employee.lastName}`}
                        onError={(e) => {
                          (e.currentTarget as HTMLImageElement).src = userImg;
                        }}
                      />
                
                    </td>
                    {dynamicColumns.map((col) => {
                      const value = employee[col.key];

                      // hide SPsubDepartment if it contains "*"
                      if (col.key === "SPsubDepartment" && typeof value === "string" && value.includes("*")) {
                        return <td key={col.key} style={col.style}></td>;
                      }

                      return (
                        <td key={col.key} style={col.style}>
                          {col.key === "Email" && value ? (
                            <a href={`mailto:${value}`}>{value}</a>
                          ) : col.key === "SPMobilePhone" && value ? (
                            <a href={`tel:${value}`}>{value}</a>
                          ) : typeof value === "boolean" ? (
                            value ? "כן" : "לא"
                          ) : isValidDateValue(value) ? (
                            formatDate(value as string)
                          ) : (
                            value || ""
                          )}
                        </td>

                      );
                    })}

                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>

        <div className={styles.pagination}>
          <button
            className={styles.paginationInfo}
            onClick={() => handlePageChange(totalPages)}
            disabled={currentPage === totalPages}
          >
            אחרון
          </button>
          {renderPagination()}
          <button
            className={styles.paginationInfo}
            onClick={() => handlePageChange(1)}
            disabled={currentPage === 1}
          >
            ראשון
          </button>
        </div>

        {isModalOpen && selectedEmployee && (
          <div className={styles.modalOverlay} onClick={() => setIsModalOpen(false)}>
            <div className={styles.modalPopup} onClick={(e) => e.stopPropagation()}>
              <div className={styles.modalContentBox}>
                <div className={styles.framedAvatarWrapper}>
                  {selectedEmployee && (
                    <img
                      src={getImageUrl(selectedEmployee)}
                      className={styles.avatarFrameCard}
                      alt="Frame"
                      onError={(e) => {
                        (e.currentTarget as HTMLImageElement).src = userImg;
                      }}
                    />
                  )}
                </div>

                <div className={styles.modalTextContent}>
                  <div
                    className={styles.topModalTextContent}
                    dir="rtl" // Forces right-to-left layout
                  >
                    <div className={styles.popupHeaderContiner}>
                      {/* Heading (Fullname) */}
                      {selectedEmployee?.spFullName && (
                        <span className={styles.spFullName}>
                          {selectedEmployee.spFullName}
                        </span>
                      )}

                      {/* Subheading (spFullName) */}
                      {selectedEmployee?.Title && (
                        <span className={styles.linkTitle}>
                          {selectedEmployee.Title}
                        </span>
                      )}
                    </div>


                    {/* Determine card columns dynamically per tab */}
                    {(
                      activeTab === "עובדי קשת"
                        ? props.selectedColumnsCardFirstTab
                        : activeTab === "עובדי הפקות"
                          ? props.selectedColumnsCardSecondTab
                          : activeTab === "KI"
                            ? props.selectedColumnsCardThirdTab
                            : []
                    )?.map((colKey: string) => {
                      const columnInfo = props.listColumns?.find(c => c.key === colKey);
                      const value = selectedEmployee[colKey];

                      // Skip rendering if empty
                      if (value === null || value === undefined || value === '' || (Array.isArray(value) && value.length === 0)) {
                        return null;
                      }

                      return (
                        <>
                          {/* Skip rendering if SPsubDepartment has "*" */}
                          {!(colKey === "SPsubDepartment" && typeof value === "string" && value.includes("*")) && (
                            <p key={colKey} className={styles.topModalTextContentField}>
                              <span className={styles.label}>
                                {columnInfo ? columnInfo.text : colKey}:
                              </span>{" "}
                              <strong
                                className={`${styles.strongText} ${colKey === "Department" ? styles.highlight : ""}`}
                              >
                                {colKey === "SPmanager" && value ? (
                                  value
                                ) : colKey === "Email" && value ? (
                                  <a href={`mailto:${value}`}>{value}</a>
                                ) : typeof value === "boolean" ? (
                                  value ? "כן" : "לא"
                                ) : isValidDateValue(value) ? (
                                  formatDate(value as string)
                                ) : (
                                  value
                                )
                                }
                              </strong>
                            </p>
                          )}
                        </>
                      );
                    })}
                  </div>

                  <div className={styles.layer1Wrapper}>
                    <img src={layerImg} className={styles.avatarFrame} alt="Layer" />
                  </div>
                </div>


                <button className={styles.closeButton} onClick={() => setIsModalOpen(false)}>
                  ×
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default KeshetNet;
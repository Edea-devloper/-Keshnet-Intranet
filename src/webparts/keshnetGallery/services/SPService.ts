import { SPFI, spfi, SPFx as spSPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/graph/users";
import "@pnp/sp/fields";
import "@pnp/sp/presets/all";


// Initialize SharePoint context
let sp: SPFI;

export const initializeSP = (context: any): void => {
  sp = spfi().using(spSPFx(context));
};

export const getSPListItemsById = async (listId: string) => {
  try {
    const today = new Date();
    const todayISO = today.toISOString();

    let filterQuery = "IsActive eq 1"; // base filter

    // Case 1: Both StartDate and EndDate are provided
    filterQuery += ` and ((StartDate ne null and EndDate ne null and StartDate le datetime'${todayISO}' and EndDate ge datetime'${todayISO}'))`;

    // Case 2: Only StartDate provided
    // filterQuery += ` or (StartDate ne null and EndDate eq null and StartDate le datetime'${todayISO}')`;

    // Case 3: Only EndDate provided
    // filterQuery += ` or (StartDate eq null and EndDate ne null and EndDate ge datetime'${todayISO}'))`;

    const result = await sp.web.lists
      .getById(listId)
      .items
      .select(
        "*",
        "Employee/Id",
        "Employee/Title",
        "Employee/EMail",
        "Employee/JobTitle",
        "AttachmentFiles/FileName",
        "AttachmentFiles/ServerRelativeUrl",
      )
      .expand("Employee,AttachmentFiles")
      .filter(filterQuery).orderBy("Created", false)
      .top(5000)();

    return result;
  } catch (error) {
    console.error("Error fetching list by ID:", error);
    return [];
  }

};

export const getAllImages = async (libraryId: string) => {
  try {
    const result = await sp.web.lists
      .getById(libraryId)
      .items
      .select(
        "Id",
        "Title",
        "FileLeafRef",          // File name
        "FileRef",              // File path
        "Created",
        "Modified",
        "Author/Id",
        "Author/Title",
        "Editor/Id",
        "Editor/Title",
        "EmployeeName/Title",
        "EmployeeName/EMail"
      )
      .expand("Author,Editor,EmployeeName")
      .top(5000)();

    return result;
  } catch (error) {
    console.error("Error fetching files:", error);
    return [];
  }
};


export const GetTemplateList = async (listId: string) => {
  try {
    const result = await sp.web.lists
      .getById(listId)
      .items
      .select("*")
      .top(5000)();

    return result;
  } catch (error) {
    console.error("Error fetching files:", error);
    return [];
  }
};

export const getEmployees = async (listId: string) => {
  try {
    listId = listId?.replace(/[{}]/g, "");

    const result = await sp.web.lists
      .getById(listId)
      .items
      .top(5000)
      .select("*")
      .expand("AttachmentFiles")();

    return result;
  } catch (error) {
    console.error("Error fetching list by ID:", error);
    return [];
  }
};
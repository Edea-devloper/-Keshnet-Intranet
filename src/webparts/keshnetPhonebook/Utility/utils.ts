import { getSP } from "./getSP";
import "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/files";
import "@pnp/sp/folders";

export const getSPListItemsById = async (listId: string, context: any) => {
  try {
    listId = listId?.replace(/[{}]/g, "");
    const _sp = getSP(context);

    const result = await _sp.web.lists
      .getById(listId)
      .items
      .top(5000)
      .select("*,Id,Title,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl")
      .expand("AttachmentFiles")();

    return result;
  } catch (error) {
    console.error("Error fetching list by ID:", error);
    return [];
  }
};

export const getAllImages = async (libraryId: string, context: any) => {
  try {
    const _sp = getSP(context);
    const result = await _sp.web.lists
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
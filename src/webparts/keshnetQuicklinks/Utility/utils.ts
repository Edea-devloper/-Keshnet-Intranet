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

export const getListItemsFormainList = async (selectedList: any): Promise<any> => {
    try {
        const _sp = getSP();

        // Get today's date in UTC format (start and end of the day)
        const today = new Date();
        const startOfDay = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()));
        const endOfDay = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate() + 1));

        // Convert to ISO strings for SharePoint filter
        const startIso = startOfDay.toISOString();
        const endIso = endOfDay.toISOString();

        // Filter where BirthDate is between start and end of today
        return _sp.web.lists
            .getById(selectedList.title) // GUID of the list
            .items
            .select('Fullname', 'BirthDate', 'UserImage', 'Attachments', 'AttachmentFiles')
            .expand('AttachmentFiles')
            .filter(`BirthDate ge datetime'${startIso}' and BirthDate lt datetime'${endIso}'`)
            .orderBy('ID', false)
            .top(4000)();
    } catch (error) {
        console.error("An error occurred while fetching list data:", error);
        throw error;
    }
};


export const getLatestQuickLink = async (selectedList: any): Promise<any> => {
    try {
        const _sp = getSP();

        const items = await _sp.web.lists
            .getById(selectedList.title)
            .items
            .select("*", "AttachmentFiles")
            .expand("AttachmentFiles")
            .filter("Active eq 1")
            .orderBy("Modified", false)
            .top(4)(); // Limit to 4 items

        return items;
    } catch (error) {
        console.error("Error fetching the latest QuickLink items:", error);
        throw error;
    }
};

export const getCurrentUser = async (): Promise<any> => {
  try {
    const _sp = getSP();

    // Get current user details
    const user = await _sp.web.currentUser();

    // user will have properties like Id, Title (Display Name), LoginName, Email
    return {
      id: user.Id,
      name: user.Title,      // Display name
      login: user.LoginName, // LoginName (e.g., i:0#.f|membership|abc@domain.com)
      email: user.Email      // User email
    };
  } catch (error) {
    console.error("Error fetching current user:", error);
    throw error;
  }
};


export const getSPListItemsById = async (listId: string, context: any) => {
  try {
    listId = listId?.replace(/[{}]/g, "");
    const _sp = getSP(context);

    const result = await _sp.web.lists
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
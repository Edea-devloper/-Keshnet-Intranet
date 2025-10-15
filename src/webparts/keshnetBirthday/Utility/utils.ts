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

        const today = new Date();
        const todayMonth = today.getMonth(); // 0-based
        const todayDate = today.getDate(); // 1-based

        const items = await _sp.web.lists
            .getById(selectedList.title)
            .items
            .select('*')
            .expand('AttachmentFiles')
            .top(4000)();

        // If no items at all
        if (!items || items.length === 0) {
            console.warn("No records found in the list.");
            return [];
        }

        // Filter items safely
        const filteredItems = items.filter(item => {

            // Check if sharepoint_status is "active"
            if (item?.sharepoint_status !== "active") return false;

            if (!item.spBirthday) return false; // Skip if spBirthday missing

            const bday = new Date(item.spBirthday);
            if (isNaN(bday.getTime())) return false; // Skip if invalid date

            return (
                bday.getMonth() === todayMonth &&
                bday.getDate() === todayDate
            );
        });

        // Optional: log if no matches
        if (filteredItems.length === 0) {
            console.info("No birthdays match today's date.");
        }

        return filteredItems;

    } catch (error) {
        console.error("An error occurred while fetching list data:", error);
        throw error;
    }
};

export const getAllImages = async (libraryId: string) => {
    try {
        const _sp = getSP();
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


import { spfi, SPFI, SPFx } from "@pnp/sp";
// import { spfi, SPFI, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/fields";
 
 
let _sp: SPFI | any = null;
 
export const getSP = (context?: any): SPFI => {
  if (context) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};
 
 
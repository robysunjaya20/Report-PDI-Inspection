export type Line = "LH" | "RH";
export type Model = "SU2ID LH" | "KS LH" | "SU2ID RH" | "KS RH";
export type InspectionItem = {id:string; model:Model|""; suffix:string; sampleQty:string; ng:string; detailNg:string};
export type Inspection = {id:string; inspector:string; inspectionDate:string; line:Line[]; lotDate:string; shift:"1"|"2"|"3"; items:InspectionItem[]; createdAt:string; fileName?:string; driveUrl?:string; downloadUrl?:string; folderUrl?:string};
export const STORAGE_DRAFT="pdi-inspection-draft-v1";
export const STORAGE_HISTORY="pdi-inspection-history-v1";
export const MODEL_PREFIX:Record<Model,string>={"SU2ID LH":"87610-I7","KS LH":"87610-I6","SU2ID RH":"87620-I7","KS RH":"87620-I6"};
export const emptyItem=():InspectionItem=>({id:crypto.randomUUID(),model:"",suffix:"",sampleQty:"",ng:"0",detailNg:""});
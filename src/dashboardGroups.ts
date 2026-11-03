export type DashboardGroup = {id:string;title:string;cards:string[]};
// Stable card IDs keep grouping independent of user-defined display names.
export const dashboardGroups:DashboardGroup[] = [
 {id:'today-pnl',title:"Today's P&L",cards:['total-combined-today','total-latest','total-fx-today']},
 {id:'today-fixed',title:"Today's fixed amount",cards:['fixed-today']},
 {id:'period',title:'Period totals',cards:['total-period','fixed-period','total-fx-period','total-combined-period']},
 {id:'performance',title:'Daily performance',cards:['total-average','total-change']},
 {id:'rates',title:'Exchange rates',cards:['rate-today','rate-period','rate-change']},
];

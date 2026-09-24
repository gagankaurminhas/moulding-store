function main(workbook: ExcelScript.Workbook) {
 const specs=[
  ["Deliveries",["DeliveryID","Date","Customer","Address","VehicleType","Latitude","Longitude","TruckID","RouteID","StopNumber","ServiceMinutes","WeightTons","Status","Notes"]],
  ["Trucks",["TruckID","VehicleType","CapacityTons","DefaultDriver","Active","Plate"]],
  ["Drivers",["DriverID","DriverName","DefaultTruckID","Active","Phone"]],
  ["Routes",["RouteID","Date","TruckID","DriverID","StopCount","DistanceKm","EstimatedMinutes","Status","ShareURL","LastUpdated"]]
 ];
 for(const [name,headers] of specs as [string,string[]][]) {
  let ws=workbook.getWorksheet(name); if(!ws) ws=workbook.addWorksheet(name);
  ws.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);
  ws.getRangeByIndexes(0,0,1,headers.length).setValues([headers]);
  ws.getRangeByIndexes(0,0,1,headers.length).getFormat().getFont().setBold(true);
  ws.freezePanes.freezeRows(1);
 }
 const t=workbook.getWorksheet("Trucks")!;
 t.getRange("A2:F9").setValues([
 ["5T-01","5 Ton",5,"","Yes",""],["5T-02","5 Ton",5,"","Yes",""],["5T-03","5 Ton",5,"","Yes",""],["5T-04","5 Ton",5,"","Yes",""],
 ["3T-01","3 Ton",3,"","Yes",""],["3T-02","3 Ton",3,"","Yes",""],["VAN-01","Sprinter","","","Yes",""],["VAN-02","Sprinter","","","Yes",""]
 ]);
 for(const [name] of specs as [string,string[]][]) {const ws=workbook.getWorksheet(name)!;const r=ws.getUsedRange();if(r){r.getFormat().autofitColumns();if(!ws.getTables().length)ws.addTable(r,true).setName(name+"Table");}}
}
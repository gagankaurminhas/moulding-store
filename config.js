const APP_CONFIG = {
  clientId: "ed8479d2-5a79-4030-9b29-007ae8b98dcd",
  tenantId: "9b01cfb5-bb13-476b-be96-44f4139eaa56",
  redirectUri: "https://gagankaurminhas.github.io/moulding-store/",
  workbookFileName: "MouldingStoreDatabase.xlsx",
  workbookDriveId: "",
  workbookItemId: "",
  scopes: ["User.Read", "Files.ReadWrite"],
  graphBaseUrl: "https://graph.microsoft.com/v1.0",
  tables: {
    drivers: "DriversTable",
    trucks: "TrucksTable",
    orders: "OrdersTable",
    schedules: "DailyScheduleTable",
    routeStops: "RouteStopsTable",
    deliveryStatus: "DeliveryStatusTable",
    loading: "LoadingChecklistTable",
    issues: "DeliveryIssuesTable"
  },
  refreshSeconds: 15
};

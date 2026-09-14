import { Injectable } from "@nestjs/common";
import type { OrganizationContext } from "../../common/context/current-organization.decorator";

@Injectable()
export class AnalyticsService {
  getDashboardSnapshot(organization: OrganizationContext) {
    return {
      organizationId: organization.organizationId,
      currencyCode: "USD",
      period: {
        from: "2026-09-01",
        to: "2026-09-30",
      },
      metrics: {
        revenue: 0,
        expenses: 0,
        netProfit: 0,
        payroll: 0,
        assetValue: 0,
        inventoryValue: 0,
        openReceivables: 0,
        openPayables: 0,
      },
      notes: [
        "Placeholder service. Production implementation should aggregate from journal lines, assets, inventory items, contacts, invoices, and payments.",
      ],
    };
  }
}

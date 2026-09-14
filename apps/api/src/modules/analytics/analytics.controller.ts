import { Controller, Get } from "@nestjs/common";
import { CurrentOrganization, OrganizationContext } from "../../common/context/current-organization.decorator";
import { AnalyticsService } from "./analytics.service";

@Controller("analytics")
export class AnalyticsController {
  constructor(private readonly analyticsService: AnalyticsService) {}

  @Get("dashboard")
  dashboard(@CurrentOrganization() organization: OrganizationContext) {
    return this.analyticsService.getDashboardSnapshot(organization);
  }
}

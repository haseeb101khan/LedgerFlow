import { Controller, Get } from "@nestjs/common";
import { CurrentOrganization, OrganizationContext } from "../../common/context/current-organization.decorator";

@Controller("organizations")
export class OrganizationsController {
  @Get("current")
  current(@CurrentOrganization() organization: OrganizationContext) {
    return {
      organizationId: organization.organizationId,
      userId: organization.userId,
      role: organization.role,
    };
  }
}

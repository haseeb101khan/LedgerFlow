import { Body, Controller, Post } from "@nestjs/common";
import { CurrentOrganization, OrganizationContext } from "../../common/context/current-organization.decorator";
import { ImportsService, PreviewImportInput } from "./imports.service";

@Controller("imports")
export class ImportsController {
  constructor(private readonly importsService: ImportsService) {}

  @Post("preview")
  preview(@CurrentOrganization() organization: OrganizationContext, @Body() body: PreviewImportInput) {
    return this.importsService.preview(organization, body);
  }
}

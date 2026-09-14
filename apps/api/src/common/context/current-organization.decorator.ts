import { createParamDecorator, ExecutionContext } from "@nestjs/common";
import type { Request } from "express";

export type OrganizationContext = {
  organizationId: string;
  userId: string;
  role: string;
};

export const CurrentOrganization = createParamDecorator(
  (_: unknown, context: ExecutionContext): OrganizationContext => {
    const request = context.switchToHttp().getRequest<Request>();

    return {
      organizationId: String(request.headers["x-organization-id"] || ""),
      userId: String(request.headers["x-user-id"] || ""),
      role: String(request.headers["x-organization-role"] || "viewer"),
    };
  },
);

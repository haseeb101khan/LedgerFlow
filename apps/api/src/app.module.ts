import { Module } from "@nestjs/common";
import { ConfigModule } from "@nestjs/config";
import { AnalyticsModule } from "./modules/analytics/analytics.module";
import { HealthModule } from "./modules/health/health.module";
import { ImportsModule } from "./modules/imports/imports.module";
import { OrganizationsModule } from "./modules/organizations/organizations.module";

@Module({
  imports: [
    ConfigModule.forRoot({ isGlobal: true }),
    HealthModule,
    OrganizationsModule,
    ImportsModule,
    AnalyticsModule,
  ],
})
export class AppModule {}

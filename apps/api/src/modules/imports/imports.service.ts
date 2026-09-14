import { Injectable } from "@nestjs/common";
import type { OrganizationContext } from "../../common/context/current-organization.decorator";

export type PreviewImportInput = {
  targetEntity: "employees" | "assets" | "inventory_items" | "contacts" | "journal_entries";
  headers: string[];
  sampleRows: Record<string, unknown>[];
};

const canonicalFields: Record<PreviewImportInput["targetEntity"], string[]> = {
  employees: ["full_name", "department_id", "job_title", "email", "phone", "status"],
  assets: ["name", "asset_code", "category_id", "assigned_employee_id", "condition", "current_value"],
  inventory_items: ["name", "sku", "category", "quantity_on_hand", "reorder_level", "unit_cost", "selling_price"],
  contacts: ["name", "type", "email", "phone", "opening_balance"],
  journal_entries: ["entry_date", "description", "reference", "journal_lines"],
};

@Injectable()
export class ImportsService {
  preview(organization: OrganizationContext, input: PreviewImportInput) {
    const fields = canonicalFields[input.targetEntity];

    return {
      organizationId: organization.organizationId,
      targetEntity: input.targetEntity,
      detectedHeaders: input.headers,
      suggestedMapping: this.suggestMapping(input.headers, fields),
      sampleRows: input.sampleRows.slice(0, 5),
      nextStep: "stage_rows_for_validation",
    };
  }

  private suggestMapping(headers: string[], fields: string[]) {
    return Object.fromEntries(
      fields.map((field) => {
        const normalizedField = this.normalize(field);
        const match = headers.find((header) => this.normalize(header).includes(normalizedField));
        return [field, match || null];
      }),
    );
  }

  private normalize(value: string) {
    return value.toLowerCase().replace(/[^a-z0-9]/g, "");
  }
}

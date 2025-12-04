import Logger from "../utils/logger";
import type { OutputFormat } from "./types";

const RAW_TEXT = {
  format: { type: "text" },
};

const TABLE_FORMAT = {
  format: {
    type: "json_schema",
    strict: true,
    name: "TableFormat",
    schema: {
      properties: {
        data: {
          description:
            "Nested list of values for multiple spreadsheet cells. Outer list represents rows, inner list represents columns. Each element in the inner list is a cell value.",
          items: { items: { anyOf: [{ type: "number" }, { type: "string" }] }, type: "array" },
          type: "array",
        },
      },
      required: ["data"],
      title: "ExcelFormat",
      type: "object",
      additionalProperties: false,
    },
  },
};

const SINGLE_CELL_FORMAT = {
  format: {
    type: "json_schema",
    strict: true,
    name: "SingleCellFormat",
    schema: {
      properties: {
        data: {
          anyOf: [{ type: "number" }, { type: "string" }],
          description: "Single value of a spreadsheet cell",
          title: "Data",
        },
      },
      required: ["data"],
      title: "ExcelCell",
      type: "object",
      additionalProperties: false,
    },
  },
};

const formatMap: Record<OutputFormat, { format: { type: string } }> = {
  raw: RAW_TEXT,
  table: TABLE_FORMAT,
  single_cell: SINGLE_CELL_FORMAT,
};

export function getTextFormat(format: string): { format: { type: string } } {
  const textFormat = formatMap[format];
  if (!textFormat) {
    Logger.error(`Invalid format: ${format}`);
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      'Invalid format. Please use "raw", "table", or "single_cell"'
    );
  }

  return textFormat;
}

export function parseTextFormat(
  content: string,
  format: OutputFormat
): Array<Array<string | number>> {
  try {
    if (format == "table") {
      return JSON.parse(content).data;
    } else if (format == "single_cell") {
      return [[JSON.parse(content).data]];
    } else {
      return [[content]];
    }
  } catch (error) {
    // Fallback if content is not valid JSON for table and single cell format
    Logger.error("Error parsing text format:", { error, content, format });
    return [["No response content"]];
  }
}

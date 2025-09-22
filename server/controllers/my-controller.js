"use strict";
const ExcelJS = require("exceljs");

// Utility to strip HTML tags and decode common HTML entities
const stripHtml = (value) => {
  if (value === null || value === undefined) return value;
  let text = String(value);
  // Convert common line-break tags to whitespace first
  text = text.replace(/<(br|BR)\s*\/?>/g, "\n");
  text = text.replace(/<\/(p|P)>/g, "\n");
  // Remove all remaining tags
  text = text.replace(/<[^>]+>/g, "");
  // Decode common entities
  const entitiesMap = {
    "&nbsp;": " ",
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": '"',
    "&#39;": "'",
    "&#x27;": "'",
  };
  text = text.replace(/&nbsp;|&amp;|&lt;|&gt;|&quot;|&#39;|&#x27;/g, (m) => entitiesMap[m] || m);
  // Decode numeric entities
  text = text.replace(/&#(\d+);/g, (_, code) => {
    try {
      return String.fromCharCode(parseInt(code, 10));
    } catch (_) {
      return _;
    }
  });
  // Collapse whitespace
  text = text.replace(/\s+/g, " ").trim();
  return text;
};

module.exports = ({ strapi }) => ({
  async getDropDownData() {
    let excel = strapi.config.get("excel");
    let dropDownValues = [];
    let array = Object.keys(excel?.config);

    strapi?.db?.config?.models?.forEach((element) => {
      if (element?.kind == "collectionType") {
        array?.forEach((data) => {
          if (element?.uid?.startsWith(data)) {
            dropDownValues.push({
              label: element?.info?.displayName,
              value: element?.uid,
            });
          }
        });
      }
    });
    // Sort dropDownValues alphabetically by label in ascending order
    dropDownValues.sort((a, b) => a.label.localeCompare(b.label));

    return {
      data: dropDownValues,
    };
  },
  async getTableData(ctx) {
    let excel = strapi.config.get("excel");
    let uid = ctx?.query?.uid;
    let limit = ctx?.query?.limit;
    let offset = ctx?.query?.offset;
    let query = await this.restructureObject(
      excel?.config[uid],
      uid,
      limit,
      offset
    );

    let response = await strapi.db.query(uid).findMany(query);

    let header = [
      ...excel?.config[uid]?.columns,
      ...Object.keys(excel?.config[uid]?.relation),
    ];

    let where = {};

    if (excel?.config[uid]?.locale == "true") {
      where = {
        locale: "en",
      };
    }

    let count = await strapi.db.query(uid).count(where);

    let tableData = await this.restructureData(response, excel?.config[uid]);

    // Sort dropDownValues alphabetically by label in ascending order

    return {
      data: tableData,
      count: count,
      columns: header,
    };
  },
  async downloadExcel(ctx) {
    try {
      let excel = strapi.config.get("excel");

      let uid = ctx?.query?.uid;

      let query = await this.restructureObject(excel?.config[uid], uid);

      let response = await strapi.db.query(uid).findMany(query);

      let excelData = await this.restructureData(response, excel?.config[uid]);

      // Create a new workbook and add a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      // Extract column headers dynamically from the data
      let headers = [
        ...excel?.config[uid]?.columns,
        ...Object.keys(excel?.config[uid]?.relation),
      ];

      // // Transform the original headers to the desired format
      let headerRestructure = [];
      headers?.forEach((element) => {
        const formattedHeader = element
          .split("_")
          .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
          .join(" ");

        headerRestructure.push(formattedHeader);
      });

      // Define dynamic column headers
      worksheet.columns = headers.map((header, index) => ({
        header: headerRestructure[index], // Use the formatted header
        key: header,
        width: 20,
      }));

      // Define the dropdown list options for the Gender column

      // Add data to the worksheet
      excelData?.forEach((row) => {
        // Excel will provide a dropdown with these values.
        worksheet.addRow(row);
      });

      // Enable text wrapping for all columns
      worksheet.columns.forEach((column) => {
        column.alignment = { wrapText: true };
      });

      // Freeze the first row
      worksheet.views = [
        { state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A" },
      ];

      // Write the workbook to a file
      const buffer = await workbook.xlsx.writeBuffer();

      return buffer;
    } catch (error) {
      console.error("Error writing buffer:", error);
    }
  },
  async restructureObject(inputObject, uid, limit, offset) {
    let excel = strapi.config.get("excel");

    let where = {};

    if (excel?.config[uid]?.locale == "true") {
      where = {
        locale: "en",
      };
    }
    let orderBy = {
      id: "asc",
    };

    const restructuredObject = {
      select: inputObject.columns || "*",
      populate: {},
      where,
      orderBy,
      limit: limit,
      offset: offset,
    };

    for (const key in inputObject.relation) {
      restructuredObject.populate[key] = {
        select: inputObject.relation[key].column,
      };
    }

    return restructuredObject;
  },
  async restructureData(data, objectStructure) {
    const stripColumns = new Set([
      ...(objectStructure?.stripHtmlColumns || []),
      ...((objectStructure?.export && objectStructure?.export?.stripHtmlColumns) || []),
    ]);

    return data.map((item) => {
      const restructuredItem = {};

      // Restructure main data based on columns
      for (const key of objectStructure.columns) {
        if (key in item) {
          const rawValue = item[key];
          if (stripColumns.has(key)) {
            if (Array.isArray(rawValue)) {
              restructuredItem[key] = rawValue.map((v) => stripHtml(v)).join(" ");
            } else {
              restructuredItem[key] = stripHtml(rawValue);
            }
          } else {
            restructuredItem[key] = rawValue;
          }
        }
      }

      // Restructure relation data based on the specified structure
      for (const key in objectStructure.relation) {
        if (key in item) {
          const column = objectStructure.relation[key].column[0];
          if (item[key] && typeof item[key] === "object") {
            if (Array.isArray(item[key]) && item[key].length > 0) {
              const joined = item[key].map((obj) => obj[column]).join(" ");
              restructuredItem[key] = stripColumns.has(key) ? stripHtml(joined) : joined;
            } else {
              const val = item[key][column];
              restructuredItem[key] = stripColumns.has(key) ? stripHtml(val) : val;
            }
          } else {
            // Handle the case where item[key] is not an object
            restructuredItem[key] = null; // Or handle it as needed
          }
        }
      }

      return restructuredItem;
    });
  },
});

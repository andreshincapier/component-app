import { test } from "@playwright/test";
import ExcelJS from "exceljs";

test("Create and write and a xlsx file with API data", async ({ request }) => {
  let response = await request.get("/breeds", {
    params: {
      limit: "5",
    },
  });
  let parsedResponse = await JSON.parse(await response.text());
  console.log(parsedResponse);
  let cats = await parsedResponse.data;
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Cats");
  const desiredHeaders = ["breed", "country", "origin"];
  const headers = Object.keys(cats[0]);
  const filteredKeys = headers.filter((key) => desiredHeaders.includes(key));

  worksheet.addRow(filteredKeys);

  cats.forEach((row) => {
    const values = filteredKeys.map((header) => row[header]);
    worksheet.addRow(values);
  });
  await workbook.xlsx.writeFile("catsexcel.xlsx");
});


test("Read xlsx data info and store in objects", async ({ request }) => {

});
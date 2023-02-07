import {
  updateExcelTable,
} from '../loc/sharepoint.js';

async function updateExcel() {
  const values = [];
  values.push(['https://main--milo--adobecom.hlx.page/drafts/sukamat/hello-world/doc3']);
  console.log('before update call');

  //const excelPath = '/drafts/localization/projects/sukamat-loc.xlsx';
  //const excelPath = '/drafts/localization/projects/data.xlsx';
  const excelPath = '/drafts/sukamat/hello-world/data.xlsx';

  await updateExcelTable(excelPath, 'URL', values);
  console.log('after update call');
}

export {
  updateExcel,
}

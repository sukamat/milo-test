import { getDocPathFromUrl, getUrlInfo } from '../loc/utils.js';

const HLX_ADMIN_STATUS = 'https://admin.hlx.page/status';

let urlInfo;

function getHelixAdminApiUrl(urlInfo, apiBaseUri) {
  return `${apiBaseUri}/${urlInfo.owner}/${urlInfo.repo}/${urlInfo.ref}`;
}

async function getDataFileStatus(helixAdminApiUrl, sharepointProjectPath) {
  let projectFileStatusJson;
  try {
    const projectFileStatusUrl = `${helixAdminApiUrl}/?editUrl=${encodeURIComponent(sharepointProjectPath)}`;
    const projectFileStatus = await fetch(projectFileStatusUrl);
    if (projectFileStatus.ok) {
      projectFileStatusJson = await projectFileStatus.json();
    }
  } catch (error) {
    throw new Error(`Could not retrieve project file status from Helix Admin Api ${error}`);
  }
  return projectFileStatusJson;
}

async function readDataFile(dataFileUrl) {
  const resp = await fetch(dataFileUrl, { cache: 'no-store' });
  const json = await resp.json();
  if (json && json?.urls?.data) {
    return json;
  }
  return undefined;
}

async function getExcelData() {
  if (!urlInfo.isValid()) {
    throw new Error('Invalid Url Parameters');
  }

  // helix API to get the details/status of the file
  const hlxAdminStatusUrl = getHelixAdminApiUrl(urlInfo, HLX_ADMIN_STATUS);
  console.log(`hlxAdminStatusUrl: ${hlxAdminStatusUrl}`);

  // get the status of the file
  const dataFileStatus = await getDataFileStatus(hlxAdminStatusUrl, urlInfo.sp);
  if (!dataFileStatus || !dataFileStatus?.webPath) {
    throw new Error('Data File does not have valid web path');
  }
  console.log('dataFileStatus :: ');
  console.log(dataFileStatus);

  const dataFilePath = dataFileStatus.webPath;
  console.log(`dataFilePath: ${dataFilePath}`);
  const dataFileUrl = `${urlInfo.origin}${dataFilePath}`;
  console.log(`dataFileUrl: ${dataFileUrl}`);
  const dataFileName = dataFileStatus.edit.name;
  console.log(`dataFileName: ${dataFileName}`);

  const excelData = {
    url: dataFileUrl,
    path: dataFilePath,
    name: dataFileName,
    excelPath: `${dataFilePath.substring(0, dataFilePath.lastIndexOf('/'))}/${dataFileName}`,
    title: '',
    description: '',
    urls: [],
    async getJson() {
      const json = await readDataFile(dataFileUrl);
      if (!json) {
        return {};
      }
      //return json;

      this.title = json.project.data[0].title;
      this.description = json.project.data[0].description;

      const urlsData = json.urls.data;
      const urls = new Map();
      urlsData.forEach((urlRow) => {
        const url = urlRow.urls;
        const docPath = getDocPathFromUrl(url);
        urls.set(url, {
          doc: {
            filePath: docPath,
            url: url,
          },
        });
      });
      this.urls = urls;

      return json;

    }
  }

  return excelData;

}

function populateHelloPage(excelData) {
  if (!excelData) {
    throw new Error('No data available');
  }
  document.getElementById('loading').classList.add('hidden');
  document.getElementsByClassName('hello-name')[0].textContent = excelData.title;
  document.getElementsByClassName('hello-description')[0].textContent = excelData.description;
  document.getElementsByClassName('hello-urls')[0].classList.remove('hidden');

  let table = document.createElement('table');
  excelData.urls.forEach((url) => {
    console.log(url.doc.filePath);
    const tr = table.insertRow();
    tr.insertCell().appendChild(document.createTextNode(url.doc.filePath));
    tr.insertCell().appendChild(document.createTextNode(url.doc.url));
  });

  document.getElementsByClassName('hello-urls')[0].appendChild(table);

}

async function init() {

  // read the data from the URL after clicking the sidekick button
  urlInfo = getUrlInfo();
  console.log(urlInfo);

  // get path to the data file from sharepoint url
  const excelData = await getExcelData();
  console.log(`excelData.url: ${excelData.url}`);
  console.log(`excelData.path: ${excelData.path}`);
  console.log(`excelData.name: ${excelData.name}`);
  console.log(`excelData.excelPath: ${excelData.excelPath}`);

  // get JSON from the excel file
  const json = await excelData.getJson();

  console.log(json);
  console.log('urls in json');
  console.log(json.urls);

  populateHelloPage(excelData);

}


export default init;

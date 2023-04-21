import { getConfig } from '../loc/config.js';
import {
  getAuthorizedRequestOption,
} from '../loc/sharepoint.js';

async function iterateSharepointTree() {
  const { sp } = await getConfig();

  const options = getAuthorizedRequestOption({
    method: 'GET',
  });

  const res = await fetch(
    `${sp.api.excel.update.fgBaseURI}:/search(q='docx')`,
    //`${sp.api.excel.update.fgBaseURI}:/children`,
    options,
  );
  if (res.ok) {
    return res.json();
  }
  throw new Error(`Failed to add worksheet ${worksheetName} to ${excelPath}.`);

}

async function recursivelyFindAllDocxFilesInMilo() {
  const { sp } = await getConfig();
  const sharePointBaseURI = `${sp.api.excel.update.baseURI}`;
  const options = getAuthorizedRequestOption({
    method: 'GET',
  });

  let docFolders = [''];
  let docFiles = [];
  return await findAllDocxInMilo(docFiles, docFolders, sharePointBaseURI, options);
}

async function findAllDocxInMilo(docFiles, docFolders, sharePointBaseURI, options) {

  while (docFolders.length != 0) {
    const uri = `${sharePointBaseURI}${docFolders.shift()}:/children`;
    const res = await fetch(uri, options);
    if (res.ok) {
      const json = await res.json();
      const files = json.value;
      if (files) {
        for (let fileObj of files) {
          if (fileObj.folder) {
            // it is a folder
            // find a better way to get the folder path
            const folderPath = fileObj.parentReference.path.replace('/drive/root:/bacom', '') + '/' + fileObj.name;
            //console.log(fileObj);
            docFolders.push(folderPath);
          } else if (fileObj?.file?.mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            const downloadUrl = fileObj['@microsoft.graph.downloadUrl'];
            const docPath = fileObj.parentReference.path + '/' + fileObj.name;
            docFiles.push({ docDownloadUrl: downloadUrl, docPath: docPath });
          }
        };
      }
    }
  }

  return docFiles;
}


async function recursivelyFindAllDocxFilesInPink() {
  const { sp } = await getConfig();
  const sharePointBaseURI = `${sp.api.excel.update.fgBaseURI}`;
  const options = getAuthorizedRequestOption({
    method: 'GET',
  });

  return findAllDocxV2(sharePointBaseURI, options);

}

let folders = [''];
let docs = [];

async function findAllDocxV2(sharePointBaseURI, options) {
  //console.log(sharePointBaseURI);

  while (folders.length != 0) {
    const uri = `${sharePointBaseURI}${folders.shift()}:/children`;
    const res = await fetch(uri, options);
    if (res.ok) {
      const json = await res.json();
      //console.log(json);
      const files = json.value;
      if (files) {
        for (let fileObj of files) {
          if (fileObj.folder) {
            // it is a folder
            // find a better way to get the folder path
            const folderPath = fileObj.parentReference.path.replace('/drive/root:/milo-pink', '') + '/' + fileObj.name;
            //console.log(fileObj);
            folders.push(folderPath);
          } else if (fileObj?.file?.mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            const downloadUrl = fileObj['@microsoft.graph.downloadUrl'];
            const docPath = fileObj.parentReference.path + '/' + fileObj.name;
            docs.push({ docDownloadUrl: downloadUrl, docPath: docPath });
          }
        };
      }
    }
  }

  return docs;
}

export {
  iterateSharepointTree,
  recursivelyFindAllDocxFilesInMilo,
  recursivelyFindAllDocxFilesInPink,
};

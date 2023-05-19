import {
  createTag,
  getDocPathFromUrl,
  getPathFromUrl,
  getUrlInfo,
  loadingON,
  loadingOFF,
} from '../loc/utils.js';
import {
  connect as connectToSP,
  getSpFiles,
  copyFile,
  saveFile,
  getFile,
  updateExcelTable,
  copyFileAndUpdateMetadata,
} from '../loc/sharepoint.js';
import {
  getSharepointConfig,
  fetchConfigJson,
  getHelixAdminConfig,
  LOC_CONFIG
} from '../loc/config.js';
import {
  updateExcel,
} from './excel.js';
import {
  iterateSharepointTree,
  recursivelyFindAllDocxFilesInMilo,
  recursivelyFindAllDocxFilesInPink,
} from './sharepoint.js';

const HLX_ADMIN_STATUS = 'https://admin.hlx.page/status';

let urlInfo;
let decoratedConfig;
let project;
let projectDetail;

function getHelixAdminApiUrl(urlInfo, apiBaseUri) {
  return `${apiBaseUri}/${urlInfo.owner}/${urlInfo.repo}/${urlInfo.ref}`;
}

async function getProjectFileStatus(helixAdminApiUrl, sharepointProjectPath) {
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

function getPinkUrl(url) {
  return url.replace('main--milo--', 'main--milo-pink--');
}

function addOrAppendToMap(map, key, value) {
  if (map.has(key)) {
    map.get(key).push(value);
  } else {
    map.set(key, [value]);
  }
}

// Child doc is where the source doc is copied to (child folder)
function getChildDocPath(docPath) {
  const prefix = docPath.substring(0, docPath.lastIndexOf("/") + 1);
  const pageName = docPath.substring(docPath.lastIndexOf("/") + 1, docPath.length);
  return `${prefix}child/${pageName}`;
}

async function getProjectData() {
  if (!urlInfo.isValid()) {
    throw new Error('Invalid Url Parameters');
  }

  // helix API to get the details/status of the file
  const hlxAdminStatusUrl = getHelixAdminApiUrl(urlInfo, HLX_ADMIN_STATUS);
  console.log(`hlxAdminStatusUrl: ${hlxAdminStatusUrl}`);

  // get the status of the file
  const projectFileStatus = await getProjectFileStatus(hlxAdminStatusUrl, urlInfo.sp);
  if (!projectFileStatus || !projectFileStatus?.webPath) {
    throw new Error('Data File does not have valid web path');
  }
  console.log('projectFileStatus :: ');
  console.log(projectFileStatus);

  const projectPath = projectFileStatus.webPath;
  console.log(`projectPath: ${projectPath}`);
  const projectUrl = `${urlInfo.origin}${projectPath}`;
  console.log(`projectUrl: ${projectUrl}`);
  const projectName = projectFileStatus.edit.name;
  console.log(`projectName: ${projectName}`);

  project = {
    url: projectUrl,
    path: projectPath,
    name: projectName,
    excelPath: `${projectPath.substring(0, projectPath.lastIndexOf('/'))}/${projectName}`,
    sp: urlInfo.sp,
    owner: urlInfo.owner,
    repo: urlInfo.repo,
    ref: urlInfo.ref,
    title: '',
    description: '',
    async getDetail() {
      const json = await readDataFile(projectUrl);
      if (!json) {
        return {};
      }
      //return json;

      //this.title = json.project.data[0].title;
      //this.description = json.project.data[0].description;

      const urlsData = json.urls.data;
      const urls = new Map();
      const filePaths = new Map();
      urlsData.forEach((urlRow) => {
        const url = urlRow.URL;
        const docPath = getDocPathFromUrl(url);
        const childDocPath = getChildDocPath(docPath);
        //const fgDocPath = docPath;
        urls.set(url, {
          doc: {
            filePath: docPath,
            url: url,
            fg: {
              url: getPinkUrl(url),
              sp: {},
            },
          },
          childDoc: {
            filePath: childDocPath,
            url: url,
            fg: {
              url: getPinkUrl(url),
              sp: {},
            },
          }
        });
        addOrAppendToMap(filePaths, docPath, `urls|${url}|doc`);
        addOrAppendToMap(filePaths, childDocPath, `urls|${url}|childDoc`);
        //addOrAppendToMap(filePaths, fgDocPath, `urls|${url}|fgDoc`);
      });
      //this.urls = urls;
      //return json;

      projectDetail = {
        url: projectUrl,
        name: projectName,
        urls,
        filePaths,
      };

      window.projectDetail = projectDetail;
      return projectDetail;

    }
  }

  return project;

}

// function populateHelloPage(project) {
//   if (!project) {
//     throw new Error('No data available');
//   }
//   document.getElementById('loading').classList.add('hidden');
//   document.getElementsByClassName('hello-name')[0].textContent = project.title;
//   document.getElementsByClassName('hello-description')[0].textContent = project.description;
//   document.getElementsByClassName('hello-urls')[0].classList.remove('hidden');

//   let table = document.createElement('table');
//   project.urls.forEach((url) => {
//     console.log(url.doc.filePath);
//     const tr = table.insertRow();
//     tr.insertCell().appendChild(document.createTextNode(url.doc.filePath));
//     tr.insertCell().appendChild(document.createTextNode(url.doc.url));
//   });

//   document.getElementsByClassName('hello-urls')[0].appendChild(table);

// }

async function copyFilesToChildFolder() {

  function updateAndDisplayCopyStatus(copyStatus, srcPath) {
    let copyDisplayText = `Copied ${srcPath} to /child folder`;
    if (!copyStatus) {
      copyDisplayText = `Failed to copy ${srcPath} to /child folder`;
    }
    loadingON(copyDisplayText);
  }

  //loadingON('button clicked');

  async function copyFileToChild(urlInfo) {
    const status = { success: false };
    try {
      const srcPath = urlInfo?.doc?.filePath;
      loadingON(`Copying ${srcPath} to /child folder`);
      // Conflict behaviour replace for copy not supported in one drive, hence if file exists,
      // then use saveFile.
      let copySuccess = false;
      if (urlInfo?.childDoc?.sp?.status !== 200) {
        const destinationFolder = `${srcPath.substring(0, srcPath.lastIndexOf('/'))}/child`;
        copySuccess = await copyFile(srcPath, destinationFolder);
        updateAndDisplayCopyStatus(copySuccess, srcPath);
      } else {
        const file = await getFile(urlInfo.doc);
        if (file) {
          const destination = urlInfo?.childDoc?.filePath;
          if (destination) {
            const saveStatus = await saveFile(file, destination);
            if (saveStatus.success) {
              copySuccess = true;
            }
          }
        }
        updateAndDisplayCopyStatus(copySuccess, srcPath);
      }
      status.success = copySuccess;
      status.srcPath = srcPath;
      status.dstPath = `${srcPath.substring(0, srcPath.lastIndexOf('/'))}/child/${urlInfo.doc.sp.name}`;
    } catch (error) {
      // eslint-disable-next-line no-console
      console.log(`Error occurred when trying to copy to child folder ${error.message}`);
    }
    return status;
  }

  const copyStatuses = await Promise.all(
    [...projectDetail.urls].map(((valueArray) => copyFileToChild(valueArray[1]))),
  );
  const failedCopies = copyStatuses
    .filter((status) => !status.success)
    .map((status) => status?.srcPath || 'Path Info Not available');

  if (failedCopies.length > 0 /*|| failedPreviews.length > 0*/) {
    let failureMessage = failedCopies.length > 0 ? `Failed to copy ${failedCopies} to child folder` : '';
    //failureMessage = failedPreviews.length > 0 ? `${failureMessage} Failed to preview ${failedPreviews}. Kindly manually preview these files before starting the project` : '';
    loadingON(failureMessage);
  } else {
    loadingOFF();
    //await refresh();
  }

}

async function copyFilesToMiloPinkFolder() {

  function updateAndDisplayCopyStatus(copyStatus, srcPath) {
    let copyDisplayText = `Copied ${srcPath} to pink folder`;
    if (!copyStatus) {
      copyDisplayText = `Failed to copy ${srcPath} to pink folder`;
    }
    loadingON(copyDisplayText);
  }

  //loadingON('button clicked');

  async function copyFilesToMiloPink(urlInfo) {
    const status = { success: false };
    try {
      const srcPath = urlInfo?.doc?.filePath;
      loadingON(`Copying ${srcPath} to pink folder`);
      // Conflict behaviour replace for copy not supported in one drive, hence if file exists,
      // then use saveFile.
      let copySuccess = false;
      if (urlInfo?.doc?.fg?.sp?.status !== 200) {
        const destinationFolder = `${srcPath.substring(0, srcPath.lastIndexOf('/'))}`;
        copySuccess = await copyFile(srcPath, destinationFolder, undefined, true);
        updateAndDisplayCopyStatus(copySuccess, srcPath);
      } else {
        const file = await getFile(urlInfo.doc, true);
        if (file) {
          const destination = urlInfo?.doc?.filePath;
          if (destination) {
            const saveStatus = await saveFile(file, destination, true);
            if (saveStatus.success) {
              copySuccess = true;
            }
          }
        }
        updateAndDisplayCopyStatus(copySuccess, srcPath);
      }
      status.success = copySuccess;
      status.srcPath = srcPath;
      status.dstPath = srcPath;
      status.url = urlInfo.doc.url;
    } catch (error) {
      // eslint-disable-next-line no-console
      console.log(`Error occurred when trying to copy to pink folder ${error.message}`);
    }
    return status;
  }

  const copyStatuses = await Promise.all(
    [...projectDetail.urls].map(((valueArray) => copyFilesToMiloPink(valueArray[1]))),
  );
  const failedCopies = copyStatuses
    .filter((status) => !status.success)
    .map((status) => status?.srcPath || 'Path Info Not available');

  // update proejct excel 
  // const copyStatusValues = [];
  // for (const obj of copyStatuses) {
  //   console.log(`${obj.srcPath} : ${obj.success}`);
  //   copyStatusValues.push([obj.url]);
  // }
  // loadingON('Update excel with copy statues...');
  // await updateExcelTable(project.excelPath, 'URL', copyStatusValues);
  // loadingON('Updated excel with copy statues...');

  if (failedCopies.length > 0 /*|| failedPreviews.length > 0*/) {
    let failureMessage = failedCopies.length > 0 ? `Failed to copy ${failedCopies} to child folder` : '';
    //failureMessage = failedPreviews.length > 0 ? `${failureMessage} Failed to preview ${failedPreviews}. Kindly manually preview these files before starting the project` : '';
    loadingON(failureMessage);
  } else {
    loadingOFF();
    //await refresh();
  }

}

async function updateExcelFile() {
  updateExcel();
}


async function iterateTree() {
  console.log('iterate tree button clicked');
  //const res = await iterateSharepointTree();

  const start = new Date();
  console.log(start);
  console.log('start iteration :: ' + start);
  const res = await recursivelyFindAllDocxFilesInMilo();
  //const res = await recursivelyFindAllDocxFilesInPink();
  const end = new Date();
  console.log('end iteration :: ' + end);
  console.log(`${(end - start) / 1000} seconds`);
  console.log('output ::');
  console.log(res);


  // let something = [];
  // something.push({ a: 'a', b: 'b' });
  // something.push({ a: 'v', b: 'f' });
  // something.push({ a: 'f', b: 'e' });
  // console.log(something);

}

async function copyFolders() {
  await copyFile('/drafts/sukamat/fg-test/folder', '/drafts/sukamat/fg-test', undefined, true);
}

let count = 0;
let pause = false;
const DELAY_TIME = 5000; //5s

async function copySpFiles() {
  async function copyFiles(filePath) {
    const status = { success: false };
    console.log('copy started');
    const copySuccess = await copyFile(filePath, '/drafts/sukamat/copy-to', undefined, false);
    console.log('copy done');
    status.success = copySuccess;
    status.path = filePath;
    return status;
  }

  const filePaths = ['/drafts/sukamat/copy-test/doc1.docx', '/drafts/sukamat/copy-test/doc2.docx'];
  // const copyStatuses = await Promise.all(
  //   filePaths.map((filePath) => copyFiles(filePath)),
  // );

  let copyStatuses;
  for (let index = 0; index < filePaths.length; index += 1) {
    count += 1;
    if (count % 2 === 0) pause = true;
    while (pause) {
      // eslint-disable-next-line no-await-in-loop
      await new Promise((res) => setTimeout(res, DELAY_TIME)).then(pause = false);
      console.log(`waiting for ${DELAY_TIME / 1000} seconds`);
    }
    // eslint-disable-next-line no-await-in-loop
    copyStatuses = await new Promise().then(copyFiles(filePaths[index]));
  }
  console.log(copyStatuses);
}

async function createFilesForStressTest() {
  console.log('STARTED: creating multiple files for stress test');
  for (let i = 501; i <= 1000; i += 1) {
    // eslint-disable-next-line no-await-in-loop
    await copyFile('/drafts/sukamat/stress-test-500/doc1.docx', '/drafts/sukamat/stress-test-1000', `doc${i}.docx`, false);
  }
  console.log('COMPLETE: creating multiple files for stress test');
}

async function createVersionForDoc() {
  console.log('update metadata');
  await copyFileAndUpdateMetadata('/drafts/sukamat/stress-test-500/doc1.docx', '/drafts/sukamat/copy-test');
  console.log('update metadata complete');
}

function setListeners() {
  document.querySelector('#copyFiles button').addEventListener('click', copyFilesToChildFolder);
  document.querySelector('#copyFilesToPink button').addEventListener('click', copyFilesToMiloPinkFolder);
  document.querySelector('#updateExcel button').addEventListener('click', updateExcelFile);
  document.querySelector('#iterateTree button').addEventListener('click', iterateTree);
  document.querySelector('#copyFolders button').addEventListener('click', copyFolders);
  document.querySelector('#copySpFiles button').addEventListener('click', copySpFiles);
  document.querySelector('#createFilesForStressTest button').addEventListener('click', createFilesForStressTest);
  document.querySelector('#createVersionForDoc button').addEventListener('click', createVersionForDoc);
  document.querySelector('#loading').addEventListener('click', loadingOFF);
}

async function getConfig() {
  if (!decoratedConfig) {
    urlInfo = getUrlInfo();
    if (urlInfo.isValid()) {
      const configPath = `${urlInfo.origin}${LOC_CONFIG}`;
      const configJson = await fetchConfigJson(configPath)
      decoratedConfig = {
        sp: getSharepointConfig(configJson),
        admin: getHelixAdminConfig(),
      }
    }
  }
  return decoratedConfig;
}

function setProjectUrl() {
  const projectName = project.name.replace(/\.[^/.]+$/, '').replaceAll('_', ' ');
  document.getElementById('project-url').innerHTML = `<a href="${project.sp}">${projectName}</a>`;
}

function getProjectDetailContainer() {
  const container = document.getElementsByClassName('project-detail')[0];
  container.innerHTML = '';
  return container;
}

function createRow(classValue = 'default') {
  return createTag('tr', { class: `${classValue}` });
}

function createColumn(innerHtml, classValue = 'default') {
  const $th = createTag('th', { class: `${classValue}` });
  if (innerHtml) {
    $th.innerHTML = innerHtml;
  }
  return $th;
}

function createHeaderColumn(innerHtml) {
  return createColumn(innerHtml, 'header');
}

async function createTableWithHeaders(config) {
  const $table = createTag('table');
  const $tr = createRow('header');
  $tr.appendChild(createHeaderColumn('URL'));
  $tr.appendChild(createHeaderColumn('Source File'));
  $tr.appendChild(createHeaderColumn('Child Folder Copy'));
  $tr.appendChild(createHeaderColumn('Child Page Info'));
  $tr.appendChild(createHeaderColumn('Pink Folder Copy'));
  $tr.appendChild(createHeaderColumn('Pink Page Info'));
  //$tr.appendChild(createHeaderColumn('En Langstore File'));  
  //await appendLanguages($tr, config, projectDetail.englishCopyProjects, 'English Copy');
  //await appendLanguages($tr, config, projectDetail.rolloutProjects, 'Rollout');
  //await appendLanguages($tr, config, projectDetail.translationProjects);
  $table.appendChild($tr);
  return $table;
}

function getAnchorHtml(url, text) {
  return `<a href="${url}" target="_new">${text}</a>`;
}

function getLinkedPagePath(spShareUrl, pagePath) {
  return getAnchorHtml(spShareUrl.replace('<relativePath>', pagePath), pagePath);
}

function getLinkOrDisplayText(spViewUrl, docStatus) {
  const pathOrMsg = docStatus.msg;
  return docStatus.hasSourceFile ? getLinkedPagePath(spViewUrl, pathOrMsg) : pathOrMsg;
}

function getSharepointStatus(doc, isPink) {
  let sharepointStatus = 'Connect to Sharepoint';
  let hasSourceFile = false;
  let modificationInfo = 'N/A';
  if (!isPink && doc && doc.sp) {
    if (doc.sp.status === 200) {
      sharepointStatus = `${doc.filePath}`;
      hasSourceFile = true;
      modificationInfo = `By ${doc.sp?.lastModifiedBy?.user?.displayName} at ${doc.sp?.lastModifiedDateTime}`;
    } else {
      sharepointStatus = 'Source file not found!';
    }
  } else {
    if (doc.fg.sp.status === 200) {
      sharepointStatus = `${doc.filePath}`;
      hasSourceFile = true;
      modificationInfo = `By ${doc.fg.sp?.lastModifiedBy?.user?.displayName} at ${doc.fg.sp?.lastModifiedDateTime}`;
    } else {
      sharepointStatus = 'Source file not found!';
    }
  }
  return { hasSourceFile, msg: sharepointStatus, modificationInfo };
}

function showButtons(buttonIds) {
  buttonIds.forEach((buttonId) => {
    document.getElementById(buttonId).classList.remove('hidden');
  });
}

async function displayProjectDetail() {
  if (!projectDetail) {
    return;
  }
  const config = await getConfig();
  if (!config) {
    return;
  }
  const container = getProjectDetailContainer();

  // TODO: Refer displayProjectDetail() in loc ui.js
  // Need to create a table and add the URL information

  const $table = await createTableWithHeaders(config);
  //const spViewUrl = await getSpViewUrl();
  const spViewUrl = config.sp.shareUrl;
  const fgSpViewUrl = config.sp.fgShareUrl;

  projectDetail.urls.forEach((urlInfo, url) => {
    const $tr = createRow();
    const pageUrl = getAnchorHtml(url, getPathFromUrl(url));
    $tr.appendChild(createColumn(pageUrl));
    const usEnDocStatus = getSharepointStatus(urlInfo.doc);
    const usEnDocDisplayText = getLinkOrDisplayText(spViewUrl, usEnDocStatus);
    $tr.appendChild(createColumn(usEnDocDisplayText));
    //const langstoreDocStatus = getSharepointStatus(urlInfo.langstoreDoc);
    //const langstoreEnDisplayText = getLinkOrDisplayText(spViewUrl, langstoreDocStatus);
    //const langstoreDocExists = langstoreDocStatus.hasSourceFile;
    //$tr.appendChild(createColumn(langstoreEnDisplayText));
    //$tr.appendChild(createColumn(langstoreDocStatus.modificationInfo));
    //displayPageStatuses(url, subprojects, langstoreDocExists, $tr);

    const childDocStatus = getSharepointStatus(urlInfo.childDoc);
    const childDocDisplayText = getLinkOrDisplayText(spViewUrl, childDocStatus);
    const childDocExists = childDocStatus.hasSourceFile;
    $tr.appendChild(createColumn(childDocDisplayText));
    $tr.appendChild(createColumn(childDocStatus.modificationInfo));
    //displayPageStatuses(url, subprojects, childDocExists, $tr);

    const fgDocStatus = getSharepointStatus(urlInfo.doc, true);
    const fgDocDisplayText = getLinkOrDisplayText(fgSpViewUrl, fgDocStatus);
    const fgDocExists = fgDocStatus.hasSourceFile;
    $tr.appendChild(createColumn(fgDocDisplayText));
    $tr.appendChild(createColumn(fgDocStatus.modificationInfo));

    $table.appendChild($tr);
  });

  // const finalRow = createRow();
  // while (metdataColumns > 0) {
  //   finalRow.appendChild(createColumn());
  //   metdataColumns -= 1;
  // }
  // displayProjectStatuses(subprojects, finalRow);
  // $table.appendChild(finalRow);


  // -- DISABLING FOR DEMO
  // container.appendChild($table);
  // -- DISABLING FOR DEMO


  // let hideIds = ['send', 'reload', 'updateFragments', 'copyToEn'];
  // let showIds = projectDetail.translationProjects.size > 0 ? ['refresh'] : [];
  // const { projectStarted } = projectDetail;
  // if (!projectStarted) {
  //   showIds = ['reload', 'updateFragments', 'copyToEn'];
  //   hideIds = ['refresh'];
  //   if (connectedToGLaaS) {
  //     showIds.push('send');
  //   }
  // }

  //const showIds = ['copyFiles', 'copyFilesToPink', 'updateExcel', 'iterateTree', 'copyFolders'];
  //const showIds = ['iterateTree', 'copyFolders'];
  const showIds = ['copyFolders', 'copySpFiles', 'createFilesForStressTest', 'iterateTree', 'createVersionForDoc'];
  showButtons(showIds);
  //hideButtons(hideIds);
}


// this function is same as the one in loc project.js 
// except the usage of filePaths object in projectDetail
// ideally it can be reused from loc
async function updateProjectWithDocs(projectDetail) {
  if (!projectDetail || !projectDetail?.filePaths) {
    return;
  }
  const { filePaths } = projectDetail;
  const docPaths = [...filePaths.keys()];
  console.log(docPaths);
  const spBatchFiles = await getSpFiles(docPaths);
  spBatchFiles.forEach((spFiles) => {
    if (spFiles && spFiles.responses) {
      spFiles.responses.forEach((file) => {
        const filePath = docPaths[file.id];
        const spFileStatus = file.status;
        const fileBody = spFileStatus === 200 ? file.body : {};
        const referencePositions = filePaths.get(filePath);
        referencePositions.forEach((referencePosition) => {
          const keys = referencePosition.split('|');
          if (keys && keys.length > 0) {
            let position = projectDetail;
            keys.forEach((key) => {
              position = position[key] || position.get(key);
            });
            position.sp = fileBody;
            position.sp.status = spFileStatus;
          }
        });
      });
    }
  });

  const fgSpBatchFiles = await getSpFiles(docPaths, true);
  fgSpBatchFiles.forEach((spFiles) => {
    if (spFiles && spFiles.responses) {
      spFiles.responses.forEach((file) => {
        const filePath = docPaths[file.id];
        const spFileStatus = file.status;
        const fileBody = spFileStatus === 200 ? file.body : {};
        const referencePositions = filePaths.get(filePath);
        referencePositions.forEach((referencePosition) => {
          const keys = referencePosition.split('|');
          if (keys && keys.length > 0) {
            let position = projectDetail;
            keys.forEach((key) => {
              position = position[key] || position.get(key);
            });
            position.fg.sp = fileBody;
            position.fg.sp.status = spFileStatus;
          }
        });
      });
    }
  });
}


async function init() {

  try {
    setListeners();
    // loadingON('Fetching Localization Config...');
    // const config = await getConfig();
    // if (!config) {
    //   return;
    // }
    // loadingON('Localization Config loaded...');

    loadingON('Fetching Config...');
    const config = await getConfig();
    if (!config) {
      return;
    }
    loadingON('Config loaded...');



    // loadingON('Fetching Project Config...');
    // project = await initProject();

    loadingON('Fetching Project Data ...');
    project = await getProjectData();
    console.log(`project.url: ${project.url}`);
    console.log(`project.path: ${project.path}`);
    console.log(`project.name: ${project.name}`);
    console.log(`project.excelPath: ${project.excelPath}`);


    // loadingON('Refreshing Project Config...');
    // await project.purge();
    // loadingON('Fetching Project Config after refresh...');
    // await fetchProjectFile(project.url, 1);
    // project = await initProject();
    // if (!project) {
    //   loadingON('Could load project file...');
    //   return;
    // }

    loadingON(`Fetching project details for ${project.url}`);
    setProjectUrl();
    projectDetail = await project.getDetail();
    console.log('projectDetail:: ');
    console.log(projectDetail);
    loadingON('Project Details loaded...');


    loadingON('Connecting now to Sharepoint...');
    const connectedToSp = await connectToSP();
    if (!connectedToSp) {
      loadingON('Could not connect to sharepoint...');
      return;
    }


    loadingON('Connected to Sharepoint! Updating the Sharepoint Status...');
    await updateProjectWithDocs(projectDetail);

    // loadingON('Update Rollout Projects...');
    // await handleRolloutProjects();
    // if (projectDetail?.translationProjects.size > 0) {
    //   loadingON('Connecting now to GLaaS...');
    //   await connectToGLaaS(async () => {
    //     loadingON('Connected to GLaaS! Updating the GLaaS Status...');
    //     await updateGLaaSStatus(projectDetail, async () => {
    //       loadingON('Status updated! Updating UI..');
    //       await displayProjectDetail();
    //       loadingOFF();
    //     });
    //   });
    // } else {
    //   await displayProjectDetail();
    // }
    // loadingON('App loaded..');
    // loadingOFF();
    // ----- END: NOT SURE -----

    loadingON('Updating UI..');
    await displayProjectDetail();
    loadingON('UI updated..');
    loadingOFF();


  } catch (error) {
    loadingON(`Error occurred when initializing the project ${error.message}`);
  }

}


export default init;

import {
  createTag,
  getDocPathFromUrl,
  getPathFromUrl,
  getUrlInfo,
  loadingON,
  loadingOFF
} from '../loc/utils.js';
import {
  connect as connectToSP,
  getSpFiles,
  copyFile,
  updateExcelTable,
} from '../loc/sharepoint.js';
import {
  getSharepointConfig,
  fetchConfigJson,
  getHelixAdminConfig,
  LOC_CONFIG
} from '../loc/config.js';

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

function addOrAppendToMap(map, key, value) {
  if (map.has(key)) {
    map.get(key).push(value);
  } else {
    map.set(key, [value]);
  }
}

// Child doc is where the source doc is copied to (child folder)
function getChildDocPath(docPath) {
  var prefix = docPath.substring(0, docPath.lastIndexOf("/") + 1);
  var pageName = docPath.substring(docPath.lastIndexOf("/") + 1, docPath.length);
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

      this.title = json.project.data[0].title;
      this.description = json.project.data[0].description;

      const urlsData = json.urls.data;
      const urls = new Map();
      const filePaths = new Map();
      urlsData.forEach((urlRow) => {
        const url = urlRow.urls;
        const docPath = getDocPathFromUrl(url);
        const childDocPath = getChildDocPath(docPath);
        urls.set(url, {
          doc: { filePath: docPath },
          childDoc: { filePath: childDocPath },
        });
        addOrAppendToMap(filePaths, docPath, `urls|${url}|doc`);
        addOrAppendToMap(filePaths, childDocPath, `urls|${url}|childDoc`);
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

function setListeners() {
  document.querySelector('#copyFiles button').addEventListener('click', copyFilesToChildFolder);
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
  $tr.appendChild(createHeaderColumn('Destination File'));
  //$tr.appendChild(createHeaderColumn('En Langstore File'));
  $tr.appendChild(createHeaderColumn('Copied Page Info'));
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

function getSharepointStatus(doc) {
  let sharepointStatus = 'Connect to Sharepoint';
  let hasSourceFile = false;
  let modificationInfo = 'N/A';
  if (doc && doc.sp) {
    if (doc.sp.status === 200) {
      sharepointStatus = `${doc.filePath}`;
      hasSourceFile = true;
      modificationInfo = `By ${doc.sp?.lastModifiedBy?.user?.displayName} at ${doc.sp?.lastModifiedDateTime}`;
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

    $table.appendChild($tr);
  });

  // const finalRow = createRow();
  // while (metdataColumns > 0) {
  //   finalRow.appendChild(createColumn());
  //   metdataColumns -= 1;
  // }
  // displayProjectStatuses(subprojects, finalRow);
  // $table.appendChild(finalRow);

  container.appendChild($table);

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

  const showIds = ['copyFiles'];
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

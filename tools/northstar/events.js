import parseMarkdown from '../loc/helix/parseMarkdown.bundle.js';
import { mdast2docx } from '../loc/helix/mdast2docx.bundle.js';

async function mdastToDocx() {
  const page = 'http://localhost:3000/drafts/sukamat/hello-world/doc1';
  const response = await fetch(`${page}.md`);
  const resText = await response.text();
  const state = { content: { data: resText }, log: '' };
  try {
    await parseMarkdown(state);
  } catch (error) {
    console.log('Error occurred when parsing markdown', error);
  }

  const { mdast } = state.content;

  console.log("MDAST :: ");
  console.log(mdast);

  let docx = {};
  try {
    console.log('Coverting mdast to docx');
    docx = await mdast2docx(mdast);
    console.log('Covertion completed');
    console.log('docx file :: ');
    console.log(docx);
  } catch (error) {
    console.log('Error occurred when coverting mdast to docx', error);
  }

  return docx;
}

async function init() {

  try {
    const docx = await mdastToDocx();
    console.log(docx.size);

  } catch (error) {
    loadingON(`Error occurred when initializing the project ${error.message}`);
  }

}


export default init;

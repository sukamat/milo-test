export default async function init(el) {
  console.log(el);
  const p = el.querySelector('p');
  const heading = document.createElement('h2');
  heading.textContent = p.textContent;
  //el.append(heading);
  el.insertAdjacentElement('afterbegin', heading);




  const resp = await fetch('/drafts/sukamat/bar-chart.json');
  if (resp.ok) {
    const json = await resp.json();
    console.log(json.data);

    const dataHeader = document.createElement('h2');
    dataHeader.textContent = 'Data read from excel';

    json.data.forEach(metric => {
      const div = document.createElement('div');
      div.className = "metric";
      const h2 = document.createElement('h3');
      h2.textContent = metric.Browsers;

      const c = document.createElement('p');
      c.textContent = `Chrome: ${metric.Chrome}`

      const ff = document.createElement('p');
      ff.textContent = `Firefox: ${metric.Firefox}`

      div.append(dataHeader, h2, c, ff);
      el.append(div);


    });
  }
}

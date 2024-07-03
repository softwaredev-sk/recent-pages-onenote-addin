Office.onReady(function (info) {
  if (info.host === Office.HostType.OneNote) {
    // Initialize the task pane
    initializeTaskPane();
  }
});

function initializeTaskPane() {
  getCurrentPageDetails();
  displayRecentPages();

  // Register event listener for document selection changes
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onDocumentSelectionChanged);
}

function onDocumentSelectionChanged() {
  getCurrentPageDetails();
  displayRecentPages();
}

function getCurrentPageDetails() {
  OneNote.run(function (context) {
    var page = context.application.getActivePage();
    page.load("title, clientUrl, webUrl, parentSection");
    return context.sync().then(function () {
      var section = page.parentSection;
      section.load("name, notebook");
      return context.sync().then(function () {
        var notebook = section.notebook;
        notebook.load("name");
        return context.sync().then(function () {
          var breadcrumb = `${notebook.name} > ${section.name} > ${page.title}`;
          document.getElementById("breadcrumb").innerText = breadcrumb;

          // Track the current page as a recent page
          trackRecentPage(page.id, page.title, page.clientUrl, page.webUrl, section.name, notebook.name);
        });
      });
    });
  }).catch(function (error) {
    console.log("Error: " + error);
  });
}

function trackRecentPage(pageId, pageTitle, clientUrl, webUrl, sectionName, notebookName) {
  let recentPages = JSON.parse(localStorage.getItem("recentPages")) || [];

  // Remove the page if it's already in the list to update its position
  recentPages = recentPages.filter((page) => page.id !== pageId);

  // Add the page to the top of the list

  recentPages.unshift({
    id: pageId,
    title: pageTitle,
    url: clientUrl,
    wUrl: webUrl,
    section: sectionName,
    notebook: notebookName,
  });

  // Limit to 10 recent pages
  recentPages = recentPages.slice(0, 10);

  localStorage.setItem("recentPages", JSON.stringify(recentPages));
}

function displayRecentPages() {
  let recentPages = JSON.parse(localStorage.getItem("recentPages")) || [];
  let recentList = document.getElementById("recent-list");
  recentList.innerHTML = "";

  recentPages.forEach((page) => {
    let li = document.createElement("li");
    li.innerHTML = `<a href="${page.wUrl}" target="_blank">${page.title}</a><a href="${page.url}">(Open in Desktop App)</a>
    <p class="path">${page.notebook} > ${page.section}</p>`;

    recentList.appendChild(li);
  });
}

if (window.parent === window) {
  console.log("%c Note!", "font-weight: bold; font-size: 30px;color: red;");
  console.log(
    "%c This console log is a reminder that this page is hosted for onenote addin that can be sideloaded to your onenote app using manifest.xml available in this repository. The page looks empty but it works for OneNote Web",
    "font-weight: bold; font-size: 18px; margin-bottom: 12px"
  );
  console.log(
    "%c OneNote (Web) > Insert > Office Add-ins > Upload My Add-in > select manifest.xml and upload.",
    "font-weight: bold; font-size: 24px;color: yellow;"
  );

  document.querySelector(".note").textContent =
    "This page is hosted for onenote addin that can be sideloaded to your onenote app using manifest.xml available in this repository. The page looks empty but it works for OneNote Web.";
  document.querySelector(".note").appendChild(document.createElement("br"));
  const seeDocs = document.createElement("a");
  seeDocs.href = "../../index.html";
  seeDocs.textContent = "See Docs";
  document.querySelector(".note").appendChild(seeDocs);
}

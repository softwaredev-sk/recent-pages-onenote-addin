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
}

function getCurrentPageDetails() {
  let theNoblePath = "";
  OneNote.run(function (context) {
    var page = context.application.getActivePage();
    page.load("title, clientUrl, pageLevel, parentSection");
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
          trackRecentPage(page.id, page.title, page.clientUrl, section.name, notebook.name);
        });
      });
    });
  }).catch(function (error) {
    console.log("Error: " + error);
  });
}

function trackRecentPage(pageId, pageTitle, clientUrl, sectionName, notebookName) {
  let recentPages = JSON.parse(localStorage.getItem("recentPages")) || [];

  // Remove the page if it's already in the list to update its position
  recentPages = recentPages.filter((page) => page.id !== pageId);

  // Add the page to the top of the list

  recentPages.unshift({
    id: pageId,
    title: pageTitle,
    url: clientUrl,
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
    li.innerHTML = `<a href="${page.url}" style="margin: '2px 1px 2px 4px'; border-left: 2px solid black; padding: 2px; font-size: 16px">${page.title}</a>
    <div style="margin-left: 15px; font-size: 12px">${page.notebook} >${page.sectionGroup} > ${page.section}</div>`;

    recentList.appendChild(li);
  });
}

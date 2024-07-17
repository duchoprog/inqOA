document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("file-input");
  const dropArea = document.getElementById("drop-area");
  const contentForm = document.getElementById("content-form");
  const fileList = document.getElementById("file-list");
  const contentListDiv = document.getElementById("content-list");
  const finalSubmitButton = document.getElementById("final-submit");

  let section1 = [];
  let section2 = [];

  // Handle file upload via input
  fileInput.addEventListener("change", (e) => {
    handleFiles(e.target.files);
  });

  // Handle drag and drop file upload
  dropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropArea.classList.add("dragover");
  });

  dropArea.addEventListener("dragleave", () => {
    dropArea.classList.remove("dragover");
  });

  dropArea.addEventListener("drop", (e) => {
    e.preventDefault();
    dropArea.classList.remove("dragover");
    handleFiles(e.dataTransfer.files);
  });

  function handleFiles(files) {
    for (const file of files) {
      section1.push(file);
      const fileItem = document.createElement("div");
      fileItem.textContent = file.name;
      fileList.appendChild(fileItem);
    }
  }

  // Handle content form submission
  contentForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const title = document.getElementById("title-input").value;
    const content = document.getElementById("content-input").value;
    if (title && content) {
      section2.push({ title, content });
      const contentItem = document.createElement("div");
      contentItem.textContent = title;
      contentListDiv.appendChild(contentItem);
      contentForm.reset();
    }
  });

  // Handle final submission
  finalSubmitButton.addEventListener("click", () => {
    console.log(section2);
    console.log(section1);
    const resultsInput = document.getElementById("results-input").value;
    const inquiryInput = document.getElementById("inquiry-input").value;
    const data = new FormData();
    section1.forEach((file, index) => {
      data.append("files", file, file.name);
    });
    section2.forEach((content, index) => {
      data.append(`content[${index}][title]`, content.title);
      data.append(`content[${index}][content]`, content.content);
    });
    data.append("resultsPerDoc", resultsInput);
    data.append("inquiry", inquiryInput);
    console.log(data);
    // Send data to backend (example using fetch)
    fetch("/submit", {
      method: "POST",
      body: data,
    })
      .then((response) => response.json())
      .then((data) => {
        console.log(data.redirectUrl);
        if (data.success) {
          window.location.href = data.redirectUrl;
        } else {
          console.error("Failed to process the submission");
        }
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  });
});

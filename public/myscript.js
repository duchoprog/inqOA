document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("file-input");
  const previousFileInput = document.getElementById("previous-file-input");
  const dropArea = document.getElementById("drop-area");
  const previousDropArea = document.getElementById("previous-drop-area");
  const contentForm = document.getElementById("content-form");
  const fileList = document.getElementById("file-list");
  const previousFileList = document.getElementById("previous-file-list");
  const contentListDiv = document.getElementById("content-list");
  const finalSubmitButton = document.getElementById("final-submit");
  const pictureInput = document.getElementById("picture-input");

  let section1 = [];
  let section2 = [];
  let previousFile = [];

  fileInput.addEventListener("change", (e) => handleFiles(e.target.files));
  previousFileInput.addEventListener("change", (e) =>
    handlePreviousFiles(e.target.files)
  );

  dropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropArea.classList.add("dragover");
  });

  dropArea.addEventListener("dragleave", () =>
    dropArea.classList.remove("dragover")
  );

  dropArea.addEventListener("drop", (e) => {
    e.preventDefault();
    dropArea.classList.remove("dragover");
    handleFiles(e.dataTransfer.files);
  });

  previousDropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    previousDropArea.classList.add("dragover");
  });

  previousDropArea.addEventListener("dragleave", () =>
    previousDropArea.classList.remove("dragover")
  );

  previousDropArea.addEventListener("drop", (e) => {
    e.preventDefault();
    previousDropArea.classList.remove("dragover");
    handlePreviousFiles(e.dataTransfer.files);
  });

  function handleFiles(files) {
    for (const file of files) {
      section1.push(file);
      const fileItem = document.createElement("div");
      fileItem.textContent = file.name;
      fileList.appendChild(fileItem);
    }
  }

  function handlePreviousFiles(files) {
    for (const file of files) {
      previousFile.push(file);
      const fileItem = document.createElement("div");
      fileItem.textContent = file.name;
      previousFileList.appendChild(fileItem);
    }
  }

  contentForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const title = document.getElementById("title-input").value;
    const content = document.getElementById("content-input").value;
    const pictures = pictureInput.files;

    if (title && content) {
      const pictureFiles = await Promise.all(
        Array.from(pictures).map((file) => readFileAsBuffer(file))
      );
      section2.push({ title, content, pictures: pictureFiles });

      await createDocxFromContent(title, content, pictureFiles);

      const contentItem = document.createElement("div");
      contentItem.textContent = title;
      contentListDiv.appendChild(contentItem);

      contentForm.reset();
    }
  });

  async function createDocxFromContent(title, content, pictures) {
    const imageParagraphs = await Promise.all(
      pictures.map(async (image) => {
        const base64String = await bufferToBase64(image);
        return new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: base64String,
              transformation: { width: 300, height: 200 }, // Adjust size as needed
            }),
          ],
        });
      })
    );

    const doc = new docx.Document({
      sections: [
        {
          properties: {},
          children: [
            new docx.Paragraph({
              children: [new docx.TextRun({ text: title, bold: true })],
            }),
            ...content.split("\n").map(
              (line) =>
                new docx.Paragraph({
                  children: [new docx.TextRun(line)],
                })
            ),
            ...imageParagraphs,
          ],
        },
      ],
    });

    docx.Packer.toBlob(doc).then((blob) => {
      const file = new File([blob], `${title.replace(/\s+/g, "_")}.docx`, {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      section1.push(file);
    });
  }

  function readFileAsBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function bufferToBase64(buffer) {
    return btoa(
      new Uint8Array(buffer).reduce(
        (data, byte) => data + String.fromCharCode(byte),
        ""
      )
    );
  }

  const projectNameInput = document.getElementById("projectName-input");

  projectNameInput.addEventListener("input", () => {
    document.title = projectNameInput.value;
  });

  finalSubmitButton.addEventListener("click", () => {
      console.log("data enviada: ", new Date().toLocaleTimeString());

    document.querySelector(".spinnerContainer").style.display = "block";

    let projectName =
      document.getElementById("projectName-input").value ||
      "sin_nombre_asignado";
    const inquiryInput = document.getElementById("inquiry-input").value;
    const data = new FormData();

    section1.forEach((file) => data.append("files", file, file.name));
    previousFile.forEach((file) =>
      data.append("previousFiles", file, file.name)
    );
    /* section2.forEach((content, index) => {
      data.append(`content[${index}][title]`, content.title);
      data.append(`content[${index}][content]`, content.content);
    }); */
    data.append("projectName", projectName);
    data.append("inquiry", inquiryInput);

    fetch("/submit", {
      method: "POST",
      body: data,
    })
      .then((response) => {
          console.log("data recibida: ", new Date().toLocaleTimeString());

        if (response.ok) {
          return response.json();
        } else if (response.status === 302) {
          window.location.href = response.url;
        } else {
          throw new Error("Network response was not ok");
        }
      })
      .then((data) => {
        if (data.success) {
          window.location.href = data.redirectUrl;
        } else {
          console.error("Failed to process the submission");
        }
      })
      .catch((error) => console.error("Fetch error:", error));
  });
});

function startHealthCheck() {
  setInterval(async () => {
    try {
      const response = await fetch("/health");
      if (!response.ok) throw new Error("Server unavailable");
    } catch (error) {
      window.location.href = "/error";
    }
  }, 15000);
}

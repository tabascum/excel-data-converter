document.getElementById("uploadForm").addEventListener("submit", function (e) {
  e.preventDefault(); // Prevent form submission

  let fileInput = document.getElementById("file");
  let file = fileInput.files[0];

  if (!file) {
    alert("Please select a file.");
    return;
  }

  // Submit the form if a file is selected
  this.submit();
});

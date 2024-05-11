/* let inputs = document.querySelectorAll(".inputfile");

Array.prototype.forEach.call(inputs, function (input) {
  let label = input.nextElementSibling;
  let labelVal = label.innerHTML;

  input.addEventListener("change", function (e) {
    let fileName = "";
    if (this.files && this.files.length > 1)
      fileName = (this.getAttribute("data-multiple-caption") || "").replace(
        "{count}",
        this.files.length
      );
    else fileName = e.target.value.pop();

    if (fileName) label.querySelector("span").innerHTML = fileName;
    else label.innerHTML = labelVal;
  });
}); */

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

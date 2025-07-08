
Office.onReady(() => {
  // Office is ready
});

function loadEmailBody() {
  Office.context.mailbox.item.body.getAsync("text", function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("emailBody").value = result.value;
    } else {
      alert("Failed to load email body.");
    }
  });
}

function checkGrammar() {
  const text = document.getElementById("emailBody").value;
  const suggestionsDiv = document.getElementById("suggestions");
  suggestionsDiv.innerHTML = "Checking...";

  fetch("https://api.languagetoolplus.com/v2/check", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: "text=" + encodeURIComponent(text) + "&language=en-US"
  })
  .then(response => response.json())
  .then(data => {
    suggestionsDiv.innerHTML = "";
    if (data.matches.length === 0) {
      suggestionsDiv.innerHTML = "<p>No issues found.</p>";
    } else {
      data.matches.forEach(match => {
        const div = document.createElement("div");
        div.className = "suggestion";
        div.innerHTML = "<strong>Issue:</strong> " + match.message + "<br>" +
                        "<strong>Context:</strong> " + match.context.text + "<br>" +
                        "<strong>Suggestion:</strong> " + (match.replacements.map(r => r.value).join(", ") || "None");
        suggestionsDiv.appendChild(div);
      });
    }
  })
  .catch(error => {
    suggestionsDiv.innerHTML = "<p>Error checking grammar.</p>";
    console.error(error);
  });
}

// Initialize Office.js library
Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        // Office is ready to interact with Outlook
        // Add your code here to set up the add-in
        // For example, set the iframe source to the Forms survey URL
        setFormsIframeSource();
    }
});

function setFormsIframeSource() {
    // Replace this with the actual URL of your Microsoft Forms survey
    var surveyUrl = "https://forms.office.com/r/EsF2m4KjGS";
    var iframe = document.getElementById("formsIframe");

    if (iframe) {
        iframe.src = surveyUrl;
    }
}

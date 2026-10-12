/* =========================================
   STUDENTHELP - JAVASCRIPT
   ========================================= */

/* ---------- Mobile Navigation ---------- */

function toggleMenu() {

    const navMenu = document.getElementById("navMenu");

    if (navMenu) {
        navMenu.classList.toggle("show");
    }
}


/* ---------- Close Mobile Menu ---------- */

document.addEventListener("DOMContentLoaded", function () {

    const navMenu = document.getElementById("navMenu");

    if (navMenu) {
        const navLinks = navMenu.querySelectorAll("a");
        navLinks.forEach(function (link) {
            link.addEventListener("click", function () {
                navMenu.classList.remove("show");
            });
        });
    }

    const form = document.getElementById("contactForm");
    if (form) {
        form.addEventListener("submit", saveContactSubmission);
    }
});

function showContactForm() {
    document.querySelector(".contact-layout").hidden = false;
    document.getElementById("formPreview").hidden = true;
    document.getElementById("contactForm").elements.name.focus();
}

async function saveContactSubmission(event) {
    event.preventDefault();

    const form = event.currentTarget;
    const status = document.getElementById("submissionStatus");
    const submitButton = form.querySelector('button[type="submit"]');

    if (!validateForm()) {
        return;
    }

    if (!window.XLSX) {
        showSubmissionError(status, "The Excel workbook library did not load. Check your internet connection and reload this page.");
        return;
    }

    const values = {
        name: form.elements.name.value.trim(),
        email: form.elements.email.value.trim(),
        subject: form.elements.subject.value.trim(),
        message: form.elements.message.value.trim()
    };

    status.hidden = false;
    status.classList.remove("form-status-error");
    status.textContent = "Creating your Excel spreadsheet...";
    submitButton.disabled = true;

    try {
        downloadSubmissionWorkbook(values);

        document.getElementById("previewName").textContent = values.name;
        document.getElementById("previewEmail").textContent = values.email;
        document.getElementById("previewSubject").textContent = values.subject;
        document.getElementById("previewMessage").textContent = values.message;
        document.getElementById("savedMessage").textContent =
            "Your Excel spreadsheet has been downloaded.";
        form.reset();
        status.hidden = true;
        document.querySelector(".contact-layout").hidden = true;
        document.getElementById("formPreview").hidden = false;
        document.getElementById("formPreview").focus();
    } catch (error) {
        const reason = error instanceof Error ? error.message : String(error);
        status.textContent = "Could not create the Excel spreadsheet. " + reason;
        status.classList.add("form-status-error");
    } finally {
        submitButton.disabled = false;
    }
}

function downloadSubmissionWorkbook(values) {
    const submittedAt = new Date();
    const workbook = window.XLSX.utils.book_new();
    const worksheet = window.XLSX.utils.aoa_to_sheet([
        ["Submitted At", "Name", "Email", "Subject", "Message"],
        [
            submittedAt.toISOString(),
            values.name,
            values.email,
            values.subject,
            values.message
        ]
    ]);
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "Message");
    const output = window.XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const url = URL.createObjectURL(new Blob([output], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }));
    const link = document.createElement("a");
    const timestamp = submittedAt.toISOString().replace(/[:.]/g, "-");

    link.href = url;
    link.download = "support-message-" + timestamp + ".xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(function () {
        URL.revokeObjectURL(url);
    }, 1000);
}

function showSubmissionError(status, message) {
    status.hidden = false;
    status.classList.add("form-status-error");
    status.textContent = message;
}

/* ---------- Contact Form Validation ---------- */

function validateForm() {

    const name = document.getElementById("name");
    const email = document.getElementById("email");
    const subject = document.getElementById("subject");
    const message = document.getElementById("message");

    if (!name || !email || !subject || !message) {
        return true;
    }

    const nameValue = name.value.trim();
    const emailValue = email.value.trim();
    const subjectValue = subject.value.trim();
    const messageValue = message.value.trim();

    /* Check Name */

    if (nameValue.length < 2) {

        alert("Please enter a valid name.");

        name.focus();

        return false;
    }


    /* Check Email */

    const emailPattern =
        /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!emailPattern.test(emailValue)) {

        alert("Please enter a valid email address.");

        email.focus();

        return false;
    }


    /* Check Subject */

    if (subjectValue === "") {

        alert("Please select a subject.");

        subject.focus();

        return false;
    }


    /* Check Message */

    if (messageValue.length < 10) {

        alert("Please enter a message with at least 10 characters.");

        message.focus();

        return false;
    }


    return true;
}


/* ---------- Notice Modal ---------- */

function showNotice(title, text) {

    const modal = document.getElementById("noticeModal");
    const modalTitle = document.getElementById("modalTitle");
    const modalText = document.getElementById("modalText");

    if (!modal || !modalTitle || !modalText) {
        return;
    }

    modalTitle.textContent = title;

    modalText.textContent = text;

    modal.classList.add("show");

    document.body.style.overflow = "hidden";
}


/* ---------- Close Notice Modal ---------- */

function closeNotice(event) {

    const modal = document.getElementById("noticeModal");

    if (!modal) {
        return;
    }

    /*
       Close when:
       1. Close button is clicked
       2. User clicks the dark background
    */

    if (event && event.currentTarget === modal && event.target !== modal) {

        return;
    }

    modal.classList.remove("show");

    document.body.style.overflow = "";
}


/* ---------- Escape Key ---------- */

document.addEventListener("keydown", function (event) {

    if (event.key === "Escape") {

        const modal = document.getElementById("noticeModal");

        if (modal && modal.classList.contains("show")) {

            modal.classList.remove("show");

            document.body.style.overflow = "";

        }

    }

});

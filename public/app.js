const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "<your-client-id-here>",
        authority: "https://login.microsoftonline.com/<your-tenant-id-here>",
        redirectUri: "http://localhost:8000",
    },
});

let usersWithManager = [];
let allDepartments = [];
let allJobTitles = [];

// Login
async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.Read.All", "Directory.Read.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");
        await fetchUsersWithManager();
        await fetchDepartments();
        await fetchJobTitles();
    } catch (error) {
        console.error("Login failed:", error);
        alert("Login failed.");
    }
}

function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Fetch Users with Manager
async function fetchUsersWithManager() {
    const response = await callGraphApi(`/users?$expand=manager&$select=displayName,userPrincipalName,mail,manager,department,jobTitle`);
    usersWithManager = response.value.filter(user => user.manager);
}

// Fetch Departments
async function fetchDepartments() {
    allDepartments = [...new Set(usersWithManager.map(user => user.department).filter(Boolean))];
    populateDropdown("departmentDropdown", allDepartments.map(dep => ({ id: dep, name: dep })));
}

// Fetch Job Titles
async function fetchJobTitles() {
    allJobTitles = [...new Set(usersWithManager.map(user => user.jobTitle).filter(Boolean))];
    populateDropdown("jobTitleDropdown", allJobTitles.map(title => ({ id: title, name: title })));
}

// Populate Dropdown
function populateDropdown(dropdownId, items) {
    const dropdown = document.getElementById(dropdownId);
    dropdown.innerHTML = `<option value="">Select</option>`;
    items.forEach(item => {
        const option = document.createElement("option");
        option.value = item.id;
        option.textContent = item.name;
        dropdown.appendChild(option);
    });
}

// Search Function
function search() {
    const searchText = document.getElementById("searchBox").value.toLowerCase();
    const department = document.getElementById("departmentDropdown").value;
    const jobTitle = document.getElementById("jobTitleDropdown").value;

    const filteredUsers = usersWithManager.filter(user => {
        const matchesSearchText = searchText
            ? (user.displayName?.toLowerCase().includes(searchText) ||
               user.userPrincipalName?.toLowerCase().includes(searchText) ||
               user.mail?.toLowerCase().includes(searchText))
            : true;

        const matchesDepartment = department
            ? user.department === department
            : true;

        const matchesJobTitle = jobTitle
            ? user.jobTitle === jobTitle
            : true;

        return matchesSearchText && matchesDepartment && matchesJobTitle;
    });

    if (filteredUsers.length === 0) {
        alert("No matching results found.");
    }

    displayResults(filteredUsers);
}

// Display Results
function displayResults(users) {
    const outputBody = document.getElementById("outputBody");
    outputBody.innerHTML = users.map(user => `
        <tr>
            <td>${user.displayName || "N/A"}</td>
            <td>${user.userPrincipalName || "N/A"}</td>
            <td>${user.mail || "N/A"}</td>
            <td>${user.manager?.displayName || "N/A"}</td>
            <td>${user.department || "N/A"}</td>
            <td>${user.jobTitle || "N/A"}</td>
        </tr>
    `).join("");
}

// Utility Function to Call Graph API

async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.Read.All", "Directory.Read.All", "Mail.Send"],
            account,
        });

        const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
            method,
            headers: {
                Authorization: `Bearer ${tokenResponse.accessToken}`,
                "Content-Type": "application/json",
            },
            body: body ? JSON.stringify(body) : null,
        });

        if (response.ok) {
            const contentType = response.headers.get("content-type");
            if (contentType && contentType.includes("application/json")) {
                return await response.json();
            }
            return {}; // Handle responses with no body
        } else {
            const errorText = await response.text();
            console.error(`Graph API Error (${response.status}):`, errorText);
            throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error("Error in callGraphApi:", error);
        throw error;
    }
}


// Download Report as CSV
function downloadReportAsCSV() {
    const headers = ["Display Name", "UPN", "Email", "Manager", "Department", "Job Title"];
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data available to download.");
        return;
    }

    const csvContent = [headers.join(","), ...rows.map(row => row.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Users_With_Manager_Report.csv";
    link.click();
}

// Mail Report to Admin

async function sendReportAsMail() {
    const adminEmail = document.getElementById("adminEmail").value;

    if (!adminEmail) {
        alert("Please provide an admin email.");
        return;
    }

    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data to send via email.");
        return;
    }

    const emailContent = rows.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`).join("");
    const emailBody = `
        <table border="1">
            <thead>
                <tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>${emailContent}</tbody>
        </table>
    `;

    const message = {
        message: {
            subject: "Users with Manager Report",
            body: { contentType: "HTML", content: emailBody },
            toRecipients: [{ emailAddress: { address: adminEmail } }],
        },
    };

    try {
        // Ensure POST is used correctly
        await callGraphApi("/me/sendMail", "POST", message);
        alert("Report sent successfully!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report.");
    }
}

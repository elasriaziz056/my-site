const API_URL = "";

const form = document.getElementById("reportForm");
const generateBtn = document.getElementById("generateBtn");
const statusBox = document.getElementById("status");
const messageBox = document.getElementById("message");

async function checkBackend() {
    try {
        const response = await fetch(`${API_URL}/`);
        if (!response.ok) throw new Error("Backend returned an error.");

        statusBox.textContent = "Backend connected";
        statusBox.className = "status online";
    } catch (error) {
        statusBox.textContent =
            "Backend not connected. Make sure app.py is running on port 5000.";
        statusBox.className = "status offline";
    }
}

form.addEventListener("submit", async (event) => {
    event.preventDefault();

    messageBox.textContent = "";
    messageBox.className = "message";
    generateBtn.disabled = true;
    generateBtn.textContent = "Generating...";

    try {
        const data = {
            "{{utreportn°}}": document.getElementById("utReport").value,
            "{{vtreportn°}}": document.getElementById("vtReport").value,
            "{{utinspector}}": document.getElementById("utInspector").value,
            "{{vtinspector}}": document.getElementById("vtInspector").value,
            "{{date}}": document.getElementById("date").value,
            "{{stamp}}": document.getElementById("stamp").value
        };

        const count = Number(document.getElementById("count").value);

        const outputnames = [
            document.getElementById("vtFilename").value,
            document.getElementById("utFilename").value
        ];

        const response = await fetch(`${API_URL}/generate`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ data, count, outputnames })
        });

        if (!response.ok) {
            let errorMessage = "Generation failed.";
            try {
                const errorData = await response.json();
                errorMessage = errorData.error || errorMessage;
            } catch (_) {}
            throw new Error(errorMessage);
        }

        const blob = await response.blob();
        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement("a");

        link.href = downloadUrl;
        link.download = "reports.zip";
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(downloadUrl);

        messageBox.textContent =
            "Reports generated successfully. The ZIP download should start automatically.";
        messageBox.className = "message success";
    } catch (error) {
        console.error(error);
        messageBox.textContent = `Error: ${error.message}`;
        messageBox.className = "message error";
    } finally {
        generateBtn.disabled = false;
        generateBtn.textContent = "Generate Reports";
    }
});

checkBackend();

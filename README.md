# Docx MCP Server

A **FastMCP-based server** that allows ChatGPT to read, search, and modify DOCX files through the Model Context Protocol (MCP).

## Features

- 📖 **Read documents** - Full text reading and paragraph navigation
- 🔄 **Dynamic Switching** - Work on multiple documents without restarting
- 🔍 **Fuzzy search** - Find paragraphs using text matching
- ✏️ **Edit content** - Add, update, and insert paragraphs at any position
- 🎨 **Rich Formatting** - Apply bold, italic, alignment, font size, and styles
- 📊 **Table Support** - Create, read, and modify tables (markdown & tab-delimited support)
- 📝 **Template support** - Fill placeholders (`<<Name>>`, `{{Date}}`, etc.) including in tables
- ✂️ **Surgical Edits** - Find and replace text anywhere in the document
- 💾 **Save documents** - Save to new files

---

## Prerequisites

- **Python 3.10+**
- **ngrok** OR **Node.js & npm** (for localtunnel)
- **ChatGPT Pro** - With MCP Actions support

---

## Quick Start

### 1. Set Up the Environment

```bash
# Navigate to the project directory
cd "D:\MCP - Docx"

# Create virtual environment (if not already done)
python -m venv .venv

# Activate the virtual environment
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Start the MCP Server

```bash
python server.py
```

The server will start on **http://localhost:8000**.

### 3. Expose Server to Internet

You can use **localtunnel** (quick, no account) or **ngrok** (more stable).

#### Option A: localtunnel

Open a **new terminal** and run:

```bash
# Install localtunnel globally (first time only)
npm install -g localtunnel

# Start the tunnel
lt --port 8000
```

It will output a URL like:

```
your url is: https://some-random-token.loca.lt
```

#### Option B: ngrok

Open a **new terminal** and run:

```bash
ngrok http 8000
ngrok http 8000 --url hopelessly-noted-jay.ngrok-free.app
```

It will output a URL like:

```
Forwarding https://xxxx.ngrok-free.app -> http://localhost:8000
Forwarding https://hopelessly-noted-jay.ngrok-free.app -> http://localhost:8000
```

**Copy your HTTPS URL** - you'll need it for ChatGPT.

---

## Configuring ChatGPT (Developer Mode)

To use this MCP server, you need to use the **Client (Chat)** integration with the new **MCP Connectors** feature.

> **Note**: This feature requires **ChatGPT Pro/Plus** and verifying you have Developer Mode enabled.

### Step 1: Enable Developer Mode (if needed)

1. Open [ChatGPT](https://chat.openai.com/)
2. Go to **Settings** → **Connectors**
3. Check if you see a place to add connectors.
   _If not, look for "Advanced" or "Developer settings" to enable MCP Connectors._

### Step 2: Add the Custom Connector

1. In **Settings** → **Connectors**, click the **Add** or **+** button (or "Edit" -> "Add new").
2. Enter your server URL using the tunnel URL you created:

   **If using localtunnel:**

   ```
   https://some-random-token.loca.lt/sse
   ```

   **If using ngrok:**

   ```
   https://xxxx.ngrok-free.app/sse
   ```

   (Make sure to include `/sse` at the end)

3. Click **Connect** / **Save**.

### Step 3: Verify

1. Start a new chat.
2. Look for the **Connectors** icon (plug icon) or check if the "Docx Editor" tools are available.
3. Try a prompt: _"Read the document"_

---

## Usage Examples

### 📂 Managing Documents

> _"List all documents in the folder"_ > _"Switch to 'Project_Proposal.docx'"_

### 📄 Reading & Searching

> _"Read the document and summarize it"_ > _"Search for paragraphs describing the 'timeline'"_

### ✍️ Editing with Formatting

> _"Add a new paragraph at the end. Make it **bold** and center-aligned."_ > _"Insert a Heading 2 titled 'Deployment Phase' after the introduction"_ > _"Find the paragraph starting with 'Draft' and change its style to 'Quote'"_

### 📊 Working with Tables

> _"Insert a table with 3 columns: Name, Role, Email"_ > _"Read table 1 and tell me who the manager is"_ > _"Update the cell in row 2, column 3 to 'alice@example.com'"_ > _"Add a new row with: Bob, Developer, bob@example.com"_

### 🧩 Templates & Placeholders

> _"Fill in the <<ClientName>> placeholder with 'Acme Corp'"_ > _"Replace all placeholders in the document"_

### 🎯 Surgical Edits

> _"Replace 'Q1' with 'Q2' everywhere in the document"_ > _"Find 'teh' and replace it with 'the'"_

---

## Environment Variables

| Variable    | Default    | Description                                                                |
| ----------- | ---------- | -------------------------------------------------------------------------- |
| `DOCX_PATH` | `MCP.docx` | Path to the **initial** default document. You can switch files at runtime. |

**Example:**

```bash
# Start with a specific report
set DOCX_PATH=Monthly_Report.docx
python server.py
```

---

## Troubleshooting

- **"Table not detected"**: When asking GPT to insert a table, make sure it formats it clearly as markdown (`| Col | Col |`) or explicit tab-separated text.
- **Port 8000 in use**: If the server fails to start, check if another process is using port 8000.
- **Connection Refused**: Ensure your tunnel (localtunnel or ngrok) and python server are both running. Check that the URL matches exactly.
- **Url expired (localtunnel)**: localtunnel URLs can expire or change if the process restarts. Check the terminal for the new URL.

---

## Project Structure

```
D:\MCP - Docx\
├── documents/          # DOCX files for editing
│   └── MCP.docx        # Default document
├── docs/               # Reference materials
│   └── MCP.md          # OpenAI MCP guide
├── server.py           # Main FastMCP server
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

---

## License

MIT License

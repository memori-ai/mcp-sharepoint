# mcp-sharepoint

A **Model Context Protocol (MCP) server** that exposes core **Microsoft SharePoint / Microsoft Graph API** functionalities as tools usable by LLM agents (e.g. Claude Desktop).

This server is designed with an **agent-first approach**:

* no direct access to the local filesystem
* deterministic inputs and outputs

---

## ✨ Features

The server allows an agent to:

### 📁 Folder management

A path can/must be passed as an argument

* List folders
* Create new folder
* Delete folder (only if empty)
* Retrieve the folder tree structure

### 📄 Document management

A path can/must be passed as an argument

* List documents
* Read document content (PDF, Word, Excel, PowerPoint)
* Upload new document (text or binary)
* Update an existing document (replace the full content)
* Delete document

---

## 🔐 Prerequisites

* Node.js ≥ 18
* A Microsoft 365 tenant
* An app registered in **Microsoft Entra ID** with Graph permissions

### Recommended Graph permissions

```
Sites.Read.All
Sites.ReadWrite.All
Files.Read.All
Files.ReadWrite.All
```

> ⚠️ Use **Application permissions**

---

* Authentication through **Microsoft Entra ID App Registration**
* All operations are performed via **Microsoft Graph**

---

## Sharepoint guide
1. On Sharepoint, click on the left "app registration"
2. Click on top "New registration"
3. Insert a name and click "Register"
4. Open the new registered app, by clicking on its name in "app registration"
5. On top you can see the **client id** and **tenant id**
6. Click on top-right "Client credentials: 0 certificates, 0 secrets"
7. Click on "New client secret"
8. Insert the desired description and expire time, and click "Add"
9. In the table there will be the new client secret, copy the "Value" field.  
> ⚠️You can copy it only once, after you will not be able to see the "Value" again  
10. The copied "Value" is the **client secret**
11. Go to the following url: `https://graph.microsoft.com/v1.0/sites/_yourdomain_.sharepoint.com:/sites/_yoursitename_` and copy the white/plain "id" value, this is the **site id**  
> ⚠️You need to pass the access token as header param, Authorization: access_token 
12. Go to the following url: `https://_yourdomain_.sharepoint.com/sites/_yoursite_/_api/v2.0/drives` and copy the "id" value of the desired drive, this is the **drive id**
13. Go to the following url: `https://_yourdomain_.sharepoint.com/sites/_yoursite_/_api/v2.0/drives/_yourdriveid_/list` and copy the "id" value of the desired list, this is the **list id**

## ⚙️ Configuration

The authorization is obtained using **app variables**.
The agent needs to send these variables in every call.

### Required app variables

| Variable      | Description                              |
| ------------- | ---------------------------------------- |
| `tenantId`    | Microsoft Entra directory tenant ID      |
| `clientId`    | Microsoft Entra app client ID            |
| `clientSecret`| Microsoft Entra client secret            |

Other common variables needed.
These depends on the Sharepoint you are using.

| Variable      | Description                              |
| ------------- | ---------------------------------------- |
| `siteId`      | Sharepoint full site ID                  |
| `driveId`     | Sharepoint drive ID                      |
| `listId`      | Sharepoint list ID                       |

---

## 🧰 Available MCP tools

### `common parameters`

* `authParams`: {  
    clientId\<string\>,  
    clientSecret\<string\>,  
    tenantId\<string\>  
  }

* `siteDriveParams`: {  
    siteId\<string\>,  
    driveId\<string\>  
  }

### 📁 Folders

#### `getFolders`

List folders in a given path.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `path`\<string\>

---

#### `createFolder`

Create a new folder.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `path`\<string\>
* `folderName`\<string\>

---

#### `deleteFolder`

Delete a folder **only if it is empty**.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `path`\<string\>

---

#### `getFolderTree`

Retrieve the folder tree structure.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `path`\<string\> (optional, default: "root")
* `maxDepth`\<number\> (optional, default: 10)

---

### 📄 Documents

#### `getDocuments`

List documents in a folder.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `path`\<string\>

---

#### `getDocumentContent`

Retrieve and extract the textual content of a document.

Supported formats:

* Native PDFs
* Word / Excel / PowerPoint (converted to PDF via Graph)

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `filePath`\<string\>

---

#### `uploadDocument`

Upload a new document.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `filePath`\<string\>
* `content`\<string\> (text or base64)
* `contentType`\<string\> (optional, default = 'application/octet-stream')
* `overwrite`\<boolean\> (optional, default = false)

---

#### `updateDocumentContent`

Fully update an existing document.

> ⚠️ The operation replace the full content of the file.  
> ⚠️ The operation fails if the file does not exist.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `filePath`\<string\>
* `content`\<string\>
* `contentType`\<string\> (optional, default = 'application/octet-stream')

---

#### `deleteDocument`

Delete a document (moved to the SharePoint recycle bin).

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `filePath`\<string\>

---

#### `searchDocumentsByKeywords`

Search for documents containing at least one of the keywords in the given attribute.

**Input**:

* `auth`\<authParams\>
* `siteDrive`\<siteDriveParams\>
* `listId`\<string\>
* `keywords`\<Array\<string\>\>
* `attributeName`\<string\>

---

## ⚠️ Known limitations

* Partial document updates are not supported
* Large files (>4MB) require upload sessions (not yet implemented)
* Text extraction from complex PDFs/Excel files may lose structure

---

## 🔒 Security

* Never expose `client secret` publicly
* Rotate secrets regularly
* Grant the minimum required Graph permissions



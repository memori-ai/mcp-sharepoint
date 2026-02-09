import { getAccessToken } from "./authService.js";
import axios from 'axios';

/**
* Retrieve a list of folders from the specified path in SharePoint.
* @param {object} authParams - Microsoft Entra ID app authentication parameters.
* @param {object} siteDrive - SharePoint site and drive identifiers.
* @param {string} path - The path in SharePoint to retrieve folders from.
* @returns {Promise<Array>} - A promise that resolves to an array of folder objects.
*/
export async function getFolders(authParams, siteDrive, path) {
  try {
    let url = ""
    const accessToken = await getAccessToken(authParams.tenantId, authParams.clientId, authParams.clientSecret); // Ottieni il token di accesso
    if (!path  || path === "/" || path.toLowerCase() === "root") {
      url = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root/children?$filter=folder ne null`;
    } else {
      // Rimuovi eventuali slash iniziali o finali dal percorso
      const cleanPath = path.replace(/^\/|\/$/g, '');
      url = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root:/${cleanPath}:/children?$filter=folder ne null`;
    }
    const response = await axios.get(url, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Content-type": 'application/json',
      },
    });
    const folders = response.data.value || [];
    return folders;
  } catch (error) {
    console.error("Errore durante il recupero delle cartelle:", error);
    throw error;
  }
}

/**
 * Create a new folder in SharePoint at the specified path.
 * @param {object} authParams - Microsoft Entra ID app authentication parameters.
 * @param {object} siteDrive - SharePoint site and drive identifiers.
 * @param {string} path - The parent path where the folder will be created.
 * @param {string} folderName - The name of the new folder to create.
 * @returns {Promise<Object>} - A promise that resolves to the created folder object.
 */
export async function createFolder(authParams, siteDrive, path, folderName) {
  try {
    const accessToken = await getAccessToken(authParams.tenantId, authParams.clientId, authParams.clientSecret);
    let url = "";
    
    if (!path || path === "/" || path.toLowerCase() === "root") {
      url = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root/children`;
    } else {
      const cleanPath = path.replace(/^\/|\/$/g, '');
      url = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root:/${cleanPath}:/children`;
    }
    
    const response = await axios.post(
      url,
      {
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail"
      },
      {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );
    
    return response.data;
  } catch (error) {
    console.error("Errore durante la creazione della cartella:", error);
    throw error;
  }
}

/**
 * Delete an empty folder in SharePoint at the specified path.
 * @param {object} authParams - Microsoft Entra ID app authentication parameters.
 * @param {object} siteDrive - SharePoint site and drive identifiers.
 * @param {string} path - The path of the folder to delete.
 * @returns {Promise<Object>} - A promise that resolves when the folder is deleted.
 */
export async function deleteFolder(authParams, siteDrive, path) {
  try {
    const accessToken = await getAccessToken(authParams.tenantId, authParams.clientId, authParams.clientSecret);
    
    if (!path || path === "/" || path.toLowerCase() === "root") {
      throw new Error("Cannot delete root folder");
    }
    
    const cleanPath = path.replace(/^\/|\/$/g, '');
    
    // Verifica che la cartella sia vuota
    const checkUrl = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root:/${cleanPath}:/children`;
    const checkResponse = await axios.get(checkUrl, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
      },
    });
    
    if (checkResponse.data.value && checkResponse.data.value.length > 0) {
      throw new Error("Folder is not empty. Cannot delete.");
    }
    
    // Elimina la cartella
    const deleteUrl = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root:/${cleanPath}`;
    await axios.delete(deleteUrl, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
      },
    });
    
    return { success: true, message: "Folder deleted successfully" };
  } catch (error) {
    console.error("Errore durante l'eliminazione della cartella:", error);
    throw error;
  }
}

/**
 * Get a tree view of the folder structure in SharePoint.
 * @param {object} authParams - Microsoft Entra ID app authentication parameters.
 * @param {object} siteDrive - SharePoint site and drive identifiers.
 * @param {string} path - The starting path (default: "root").
 * @param {number} maxDepth - Maximum depth to traverse (default: 10).
 * @returns {Promise<Object>} - A promise that resolves to a tree structure.
 */
export async function getFolderTree(authParams, siteDrive, path = "root", maxDepth = 10) {
  try {
    const accessToken = await getAccessToken(authParams.tenantId, authParams.clientId, authParams.clientSecret);
    
    async function buildTree(currentPath, depth = 0) {
      if (depth >= maxDepth) {
        return null;
      }
      
      let url = "";
      if (!currentPath || currentPath === "/" || currentPath.toLowerCase() === "root") {
        url = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root/children?$filter=folder ne null`;
      } else {
        const cleanPath = currentPath.replace(/^\/|\/$/g, '');
        url = `https://graph.microsoft.com/v1.0/sites/${siteDrive.siteId}/drives/${siteDrive.driveId}/root:/${cleanPath}:/children?$filter=folder ne null`;
      }
      
      const response = await axios.get(url, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });
      
      const folders = response.data.value || [];
      const tree = [];
      
      for (const folder of folders) {
        const folderPath = currentPath === "root" || !currentPath 
          ? folder.name 
          : `${currentPath}/${folder.name}`;
        
        const children = await buildTree(folderPath, depth + 1);
        
        tree.push({
          name: folder.name,
          path: folderPath,
          id: folder.id,
          childCount: folder.folder.childCount,
          children: children || []
        });
      }
      
      return tree;
    }
    
    const tree = await buildTree(path);
    return tree;
  } catch (error) {
    console.error("Errore durante il recupero dell'albero delle cartelle:", error);
    throw error;
  }
}
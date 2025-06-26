"""Folder management tools for Outlook MCP Server"""
from typing import Dict, Any, Optional, List
import requests
from ..connection import get_access_token


def get_all_folders() -> Dict[str, Any]:
    """Get all mail folders"""
    try:
        access_token = get_access_token()
        url = "https://graph.microsoft.com/v1.0/me/mailFolders"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        folders = response.json().get("value", [])
        filtered_folders = [
            {
                "id": folder.get("id"),
                "displayName": folder.get("displayName"),
                "parentFolderId": folder.get("parentFolderId"),
                "childFolderCount": folder.get("childFolderCount"),
                "unreadItemCount": folder.get("unreadItemCount"),
                "totalItemCount": folder.get("totalItemCount"),
            }
            for folder in folders
        ]

        print(f"Retrieved {len(filtered_folders)} folders")
        return {"result": filtered_folders, "error": None}
        
    except Exception as e:
        error_message = f"Error getting folders: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_folder_details(folder_id: str) -> Dict[str, Any]:
    """Get details of a specific folder"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        folder = response.json()
        print(f"Retrieved folder details for: {folder.get('displayName')}")
        return {"result": folder, "error": None}
        
    except Exception as e:
        error_message = f"Error getting folder details: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def create_folder(
    display_name: str,
    parent_folder_id: Optional[str] = None
) -> Dict[str, Any]:
    """Create a new mail folder"""
    try:
        access_token = get_access_token()
        
        if parent_folder_id:
            url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{parent_folder_id}/childFolders"
        else:
            url = "https://graph.microsoft.com/v1.0/me/mailFolders"
            
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        folder_data = {
            "displayName": display_name
        }

        response = requests.post(url, headers=headers, json=folder_data, timeout=10)
        response.raise_for_status()

        folder = response.json()
        print(f"Created folder: {folder.get('displayName')}")
        return {"result": folder, "error": None}
        
    except Exception as e:
        error_message = f"Error creating folder: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def update_folder(
    folder_id: str,
    display_name: str
) -> Dict[str, Any]:
    """Update a folder's display name"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        update_data = {
            "displayName": display_name
        }

        response = requests.patch(url, headers=headers, json=update_data, timeout=10)
        response.raise_for_status()

        updated_folder = response.json()
        print(f"Updated folder: {folder_id}")
        return {"result": updated_folder, "error": None}
        
    except Exception as e:
        error_message = f"Error updating folder: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def delete_folder(folder_id: str) -> Dict[str, Any]:
    """Delete a mail folder"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.delete(url, headers=headers, timeout=10)
        response.raise_for_status()

        print(f"Deleted folder: {folder_id}")
        return {"result": "Folder deleted successfully", "error": None}
        
    except Exception as e:
        error_message = f"Error deleting folder: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_many_folders(
    folder_ids: List[str]
) -> Dict[str, Any]:
    """Get details for multiple folders"""
    try:
        access_token = get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        folders_data = []
        for folder_id in folder_ids:
            try:
                url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
                response = requests.get(url, headers=headers, timeout=10)
                response.raise_for_status()
                
                folder = response.json()
                folders_data.append({
                    "id": folder.get("id"),
                    "displayName": folder.get("displayName"),
                    "parentFolderId": folder.get("parentFolderId"),
                    "childFolderCount": folder.get("childFolderCount"),
                    "unreadItemCount": folder.get("unreadItemCount"),
                    "totalItemCount": folder.get("totalItemCount"),
                })
            except Exception as e:
                print(f"Error getting folder {folder_id}: {e}")
                folders_data.append({
                    "id": folder_id,
                    "error": str(e)
                })

        print(f"Retrieved {len(folders_data)} folder details")
        return {"result": folders_data, "error": None}
        
    except Exception as e:
        error_message = f"Error getting multiple folders: {e}"
        print(error_message)
        return {"result": None, "error": error_message}

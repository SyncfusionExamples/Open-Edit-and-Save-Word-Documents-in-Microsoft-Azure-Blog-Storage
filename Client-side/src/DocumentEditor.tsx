import React, { useRef, useState } from 'react';
import {
    DocumentEditorContainerComponent, Toolbar, CustomToolbarItemModel
} from '@syncfusion/ej2-react-documenteditor';
import AzureFileManager from './AzureFileManager.tsx';
import { ClickEventArgs } from '@syncfusion/ej2-navigations/src/toolbar/toolbar';
import { DialogUtility } from '@syncfusion/ej2-react-popups';

// Inject Document Editor toolbar dependencies
DocumentEditorContainerComponent.Inject(Toolbar);

function DocumentEditor() {
    // Backend API host URL for document operations
    const hostUrl: string = "http://localhost:62869/";
    // Reference to document editor container component
    const containerRef = useRef<DocumentEditorContainerComponent>(null);
    // Reference for dialog component
    let dialogObj: any;
    // State to hold the current document name
    const [currentDocName, setCurrentDocName] = useState<string>('None');
    // Track document modifications for auto-save functionality
    const contentChanged = React.useRef(false);

    // Custom toolbar button configuration for "New" document
    const newToolItem: CustomToolbarItemModel = {
        prefixIcon: "e-de-ctnr-new",
        tooltipText: "New",
        text: "New",
        id: "CreateNewDoc"
    };

    // Custom toolbar button configuration for opening the Azure file manager
    const openToolItem: CustomToolbarItemModel = {
        prefixIcon: "e-de-ctnr-open",
        tooltipText: "Open Azure file manager",
        text: "Open",
        id: "OpenAzureFileManager"
    };

    // Custom toolbar button configuration for downloading the document
    const downloadToolItem: CustomToolbarItemModel = {
        prefixIcon: "e-de-ctnr-download",
        tooltipText: "Download",
        text: "Download",
        id: "DownloadToLocal"
    };

    // Combined toolbar items including custom buttons and built-in features
    const toolbarItems = [newToolItem, openToolItem, downloadToolItem, 'Separator', 'Undo', 'Redo', 'Separator', 'Image', 'Table', 'Hyperlink', 'Bookmark', 'TableOfContents', 'Separator', 'Header', 'Footer', 'PageSetup', 'PageNumber', 'Break', 'InsertFootnote', 'InsertEndnote', 'Separator', 'Find', 'Separator', 'Comments', 'TrackChanges', 'Separator', 'LocalClipboard', 'RestrictEditing', 'Separator', 'FormFields', 'UpdateFields', 'ContentControl']

    // Automatically saves document to Azure storage
    const autoSaveDocument = (): void => {
        if (!containerRef.current) return;
        // Save as Blob using Docx format
        containerRef.current.documentEditor.saveAsBlob('Docx').then((blob: Blob) => {
            let exportedDocument = blob;
            let formData: FormData = new FormData();
            formData.append('documentName', containerRef.current.documentEditor.documentName);
            formData.append('data', exportedDocument);
            let req = new XMLHttpRequest();
            // Send document to backend API for Azure storage
            req.open(
                'POST',
                hostUrl + 'api/AzureFileProvider/SaveToAzure',
                true
            );
            req.onreadystatechange = () => {
                if (req.readyState === 4 && (req.status === 200 || req.status === 304)) {
                    // Auto save completed
                    // Success handler can be added here if needed
                }
            };
            req.send(formData);
        });
    };

    // Runs auto-save every second when content changes are detected
    React.useEffect(() => {
        const intervalId = setInterval(() => {
            if (contentChanged.current) {
                autoSaveDocument();
                contentChanged.current = false;
            }
        }, 1000);
        return () => clearInterval(intervalId);
    });

    // Handles document content change detection
    const handleContentChange = (): void => {
        contentChanged.current = true; // Set the ref's current value
    };

    // Handles document editor toolbar button click events
    const handleToolbarItemClick = (args: ClickEventArgs): void => {
        // Get a reference to the file manager open button
        const openButton = document.getElementById('openAzureFileManager');
        // Get the current document name from the editor
        let documentName = containerRef.current.documentEditor.documentName;
        // Remove any extension from the document name using regex
        const baseDocName = documentName.replace(/\.[^/.]+$/, '');
        // Always check if containerRef.current exists before using it
        if (!containerRef.current) return;
        switch (args.item.id) {
            case 'OpenAzureFileManager':
                // Programmatically trigger Azure file manager
                if (openButton) {
                    // save the changes before new document opening 
                    autoSaveDocument();
                    openButton.click();
                }
                break;
            case 'DownloadToLocal':
                // Initiate client-side download
                containerRef.current.documentEditor.save(baseDocName, 'Docx');
                break;
            case 'CreateNewDoc':
                // Create new document workflow
                showFileNamePrompt();
                break;
            default:
                break;
        }
    };

    // Callback function to load file selected in the file manager
    const loadFileFromFileManager = (filePath: string, fileType: string, fileName: string): void => {
        if (!containerRef.current) {
            console.error('Document Editor is not loaded yet.');
            return;
        }
        containerRef.current.documentEditor.documentName = fileName;
        // Update state with the current document name
        setCurrentDocName(fileName);
        if (fileType === '.docx' || fileType === '.doc' || fileType === '.txt' || fileType === '.rtf') {
            // Handle document files
            fetch(hostUrl + 'api/AzureFileProvider/GetDocument', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json;charset=UTF-8' },
                body: JSON.stringify({ documentName: fileName })
            })
                .then(response => {
                    if (response.status === 200 || response.status === 304) {
                        return response.json();
                    } else {
                        throw new Error('Error loading document');
                    }
                })
                .then(json => {
                    const documentEditorDiv = document.getElementById("document-editor-div")
                    if (documentEditorDiv) {
                        documentEditorDiv.style.display = "block";
                    }
                    // Open the document using the JSON data received
                    containerRef.current.documentEditor.open(JSON.stringify(json));
                })
                .catch(error => {
                    console.error('Error loading document:', error);
                });
        } else {
            alert('The selected file type is not supported for the document editor.');
        }
    };

    // List of default general document names
    const defaultFileNames = ['Untitled'];
    // Utility function to get a random default name from the list
    const getRandomDefaultName = (): string => {
        const randomIndex = Math.floor(Math.random() * defaultFileNames.length);
        return defaultFileNames[randomIndex];
    };

    //  Check if a document with a given name already exists on the Azure storage
    const validateFileExistence = async (fileName: string): Promise<boolean> => {
        try {
            const response = await fetch(hostUrl + 'api/AzureFileProvider/ValidateFileExistence', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json;charset=UTF-8' },
                body: JSON.stringify({ fileName: fileName })
            });
            if (response.ok) {
                const result = await response.json();
                return result.exists; // Backend returns { exists: boolean }
            }
            return false;
        } catch (err) {
            console.error('Error checking document existence:', err);
            return false;
        }
    };

    // Prompt dialog for entering a new document filename
    const showFileNamePrompt = (errorMessage?: string) => {
        const randomDefaultName = getRandomDefaultName();
        dialogObj = DialogUtility.confirm({
            title: 'New Document',
            width: '350px',
            content: `
                <p>Enter document name:</p> 
                <div id="errorContainer" style="color: red; margin-top: 4px;">
                ${errorMessage ? errorMessage : ''}
                </div>
                <input id="inputEle" type="text" class="e-input" value="${randomDefaultName}"/>
            `,
            okButton: { click: handleFileNamePromptOk },
            cancelButton: { click: handleFileNamePromptCancel },
        });
        // After the dialog renders, focus and select the input text.
        setTimeout(() => {
            const input = document.getElementById("inputEle") as HTMLInputElement;
            if (input) {
                input.focus();
                input.select();
            }
        }, 100);
    };

    // Handler for the OK button in the file name prompt dialog with file existence check and save 
	// The new file will be automatically saved to Azure Storage by the auto-save functionality, which is managed within the setInterval method.  
    const handleFileNamePromptOk = async () => {
        const inputElement = document.getElementById("inputEle") as HTMLInputElement;
        let userFilename = inputElement?.value.trim() || "Untitled";
        const baseFilename = `${userFilename}.docx`;

        // Check if the document already exists on the backend
        const exists = await validateFileExistence(baseFilename);
        if (exists) {
            // If the document exists, display an error message in the dialog
            const errorContainer = document.getElementById("errorContainer");
            if (errorContainer) {
                errorContainer.innerHTML = 'Document already exists. Please choose a different name.';
            }
            // Re-focus the input for correction
            if (inputElement) {
                inputElement.focus();
                inputElement.select();
            }
            return;
        }

        // Proceed with new document
        if (dialogObj) dialogObj.hide();
        containerRef.current.documentEditor.documentName = baseFilename;
        setCurrentDocName(baseFilename);
        containerRef.current.documentEditor.openBlank();
    };

    // Handler for the Cancel button in the file name prompt dialog
    const handleFileNamePromptCancel = () => {
        if (dialogObj) {
            dialogObj.hide();
        }
    };

    return (
        <div>
            <div>
                <AzureFileManager onFileSelect={loadFileFromFileManager} />
            </div>
            <div id="document-editor-div" style={{ display: "block" }}>
                <div id="document-header">
                    {currentDocName || 'None'}
                </div>
                <DocumentEditorContainerComponent
                    ref={containerRef}
                    id="container"
                    height={'650px'}
                    serviceUrl={hostUrl + 'api/AzureFileProvider/'}
                    enableToolbar={true}
                    toolbarItems={toolbarItems}
                    toolbarClick={handleToolbarItemClick}
                    contentChange={handleContentChange} // Listen to content changes
                />
            </div>
        </div>
    );
}

export default DocumentEditor;
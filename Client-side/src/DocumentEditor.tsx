import React from 'react';
import {
    DocumentEditorContainerComponent, Toolbar, CustomToolbarItemModel
} from '@syncfusion/ej2-react-documenteditor';
import AzureFileManager from './AzureFileManager.tsx';
import { ClickEventArgs } from '@syncfusion/ej2-navigations/src/toolbar/toolbar';

DocumentEditorContainerComponent.Inject(Toolbar);

function DocumentEditor() {
    const hostUrl: string = "http://localhost:62869/";
    let container: DocumentEditorContainerComponent;
    const contentChanged = React.useRef(false);
    const openToolItem: CustomToolbarItemModel = {
        prefixIcon: "e-de-ctnr-open",
        tooltipText: "Open Azure file manager",
        text: "Open",
        id: "OpenAzureFileManager"
    };

    const downloadToolItem: CustomToolbarItemModel = {
        prefixIcon: "e-de-ctnr-download",
        tooltipText: "Download",
        text: "Download",
        id: "DownloadToLocal"
    };

    const toolbarItems = ['New', openToolItem, downloadToolItem, 'Separator', 'Undo', 'Redo', 'Separator', 'Image', 'Table', 'Hyperlink', 'Bookmark', 'TableOfContents', 'Separator', 'Header', 'Footer', 'PageSetup', 'PageNumber', 'Break', 'InsertFootnote', 'InsertEndnote', 'Separator', 'Find', 'Separator', 'Comments', 'TrackChanges', 'Separator', 'LocalClipboard', 'RestrictEditing', 'Separator', 'FormFields', 'UpdateFields', 'ContentControl']

    const autoSaveDocument = (): void => {
        container.documentEditor.saveAsBlob('Docx').then((blob: Blob) => {
            let exportedDocument = blob;
            let formData: FormData = new FormData();
            formData.append('documentName', container.documentEditor.documentName);
            formData.append('data', exportedDocument);
            let req = new XMLHttpRequest();
            req.open(
                'POST',
                hostUrl + 'api/AzureFileProvider/SaveToAzure',
                true
            );
            req.onreadystatechange = () => {
                if (req.readyState === 4 && (req.status === 200 || req.status === 304)) {
                    // Auto save completed
                }
            };
            req.send(formData);
        });
    };

    React.useEffect(() => {
        const intervalId = setInterval(() => {
            if (contentChanged.current) {
                autoSaveDocument();
                contentChanged.current = false;
            }
        }, 1000);
        return () => clearInterval(intervalId);
    });

    const handleContentChange = (): void => {
        contentChanged.current = true; // Set the ref's current value
    };

    const handleToolbarClick = (args: ClickEventArgs): void => {
        const openButton = document.getElementById('openAzureFileManager');
        const documentName = container.documentEditor.documentName || 'Untitled';
        switch (args.item.id) {
            case 'OpenAzureFileManager':
                if (openButton) {
                    openButton.click();
                }
                break;
            case 'DownloadToLocal':
                container.documentEditor.save(documentName, 'Docx')
                break;
            default:
                break;
        }
    };

    // Callback function to load file selected in the file manager
    const loadFileFromFileManager = (filePath: string, fileType: string, filenName: string): void => {
        container.documentEditor.documentName = filenName;
        if (fileType === '.docx' || fileType === '.doc' || fileType === '.txt' || fileType === '.rtf') {
            // Handle document files
            fetch(hostUrl + 'api/AzureFileProvider/GetDocument', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json;charset=UTF-8' },
                body: JSON.stringify({ documentName: filenName })
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
                    container.documentEditor.open(JSON.stringify(json));
                })
                .catch(error => {
                    console.error('Error loading document:', error);
                });
        } else {
            alert('The selected file type is not supported for the document editor.');
        }
    };

    return (
        <div>
            <div>
                <AzureFileManager onFileSelect={loadFileFromFileManager} />
            </div>
            <div id="document-editor-div" style={{ display: "block" }}>
                <DocumentEditorContainerComponent
                    ref={(scope) => { container = scope; }}
                    id="container"
                    height={'590px'}
                    serviceUrl={hostUrl + 'api/AzureFileProvider/'}
                    toolbarItems={toolbarItems}
                    toolbarClick={handleToolbarClick}
                    enableToolbar={true}
                    contentChange={handleContentChange} // Listen to content changes
                />
            </div>
        </div>
    );
}

export default DocumentEditor;
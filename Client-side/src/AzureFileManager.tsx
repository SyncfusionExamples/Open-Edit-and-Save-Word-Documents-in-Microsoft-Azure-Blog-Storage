import * as React from 'react';
import { FileManagerComponent, Inject, NavigationPane, DetailsView, Toolbar } from '@syncfusion/ej2-react-filemanager';
import { DialogComponent } from '@syncfusion/ej2-react-popups';

interface AzureProps {
    onFileSelect: (filePath: string, fileType: string, fileName: string) => void; // Callback to load file in DocumentEditor
}

const AzureFileManager: React.FC<AzureProps> = ({ onFileSelect }) => {
    const hostUrl: string = "http://localhost:62869/";  // Adjust with the correct URL for file manager
    const [showFileManager, setShowFileManager] = React.useState(true);
    let fileManagerRef = React.useRef<FileManagerComponent>(null);

    // Show the file manager when open button is clicked
    const handleOpenButtonClick = () => {
        // Clear the previous selection
        if (fileManagerRef.current) {
            fileManagerRef.current.clearSelection();
        }
        setShowFileManager(true);
    };

    // Handle file open from file manager
    const handleFileOpen = (args: any) => {
        if (args.fileDetails.isFile) {
            const selectedPath = args.fileDetails.path || args.fileDetails.filterPath + args.fileDetails.name;
            const fileType = args.fileDetails.type;
            const fileName = args.fileDetails.name;
            
            onFileSelect(selectedPath, fileType, fileName); // Pass the file path and file type to load in the Document Editor
            setShowFileManager(false); // Close the File Manager Dialog
        }
    };

    return (
        <div>
            <button id = "openAzureFileManager" onClick={handleOpenButtonClick} style={{ padding: '10px 20px', cursor: 'pointer' , display: 'none'}}>
                Open Azure file manager
            </button>

            {/* File Manager Dialog */}
            <DialogComponent
                id = "dialog-component-sample"
                header="File Manager"
                visible={showFileManager}
                width="80%"
                height="80%"
                showCloseIcon={true}
                closeOnEscape={true}
                target="body"
                beforeClose={() => setShowFileManager(false)}
                onClose={() => setShowFileManager(false)} // Close the dialog when closed
            >
                <FileManagerComponent
                    id="azure-file"
                    ref={fileManagerRef}
                    ajaxSettings={{
                        url: hostUrl + 'api/AzureFileProvider/AzureFileOperations',
                        downloadUrl: hostUrl + 'api/AzureFileProvider/AzureDownload'
                    }}
                    toolbarSettings={{
                        items: ['SortBy', 'Copy', 'Paste', 'Delete', 'Refresh', 'Download', 'Selection', 'View', 'Details']
                    }}
                    contextMenuSettings={{
                        file: ['Open', 'Copy', '|', 'Delete', 'Download',  '|', 'Details'],
                        layout: ['SortBy', 'View', 'Refresh', '|', 'Paste', '|',  '|', 'Details', '|', 'SelectAll'],
                        visible: true
                    }}
                    fileOpen={handleFileOpen} // Attach the fileOpen event
                    
                >
                    <Inject services={[NavigationPane, DetailsView, Toolbar]} />
                </FileManagerComponent>
            </DialogComponent>
        </div>
    );
};

export default AzureFileManager;

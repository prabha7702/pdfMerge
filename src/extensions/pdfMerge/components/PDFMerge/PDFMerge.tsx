import * as React from 'react';
import { DefaultButton, Dialog, DialogType, Panel, PanelType, PrimaryButton, Spinner, getTheme } from 'office-ui-fabric-react';
import Document from '../Document/Document';
import { IDocument } from '../../models/IDocument';
import styles from './PDFMerge.module.scss';
import { DragDropContext, Droppable, DropResult } from 'react-beautiful-dnd';
import { SharepointService } from '../../services/SharePointService';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { PDFService } from '../../services/PDFService';
import * as strings from 'PDFMergeCommandSetStrings';

export interface IPDFMergeProps {
    listId: string;
    documents: IDocument[],
    closePanel: () => void,
    context: ListViewCommandSetContext
}

const PDFMerge: React.FC<IPDFMergeProps> = (props: IPDFMergeProps) => {
    const [documents, setDocuments] = React.useState<IDocument[]>(props.documents);
    const [filename, setFilename] = React.useState<string>(documents[0].name.split(".")[0]);
    const [isDeleteSelected, setIsDeleteSelected] = React.useState<boolean>(false);
    const [isPreviewSelected, setIsPreviewSelected] = React.useState<boolean>(false);
    const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(true);
    const [message, setMessage] = React.useState<string>("");
    const [isLoading, setIsLoading] = React.useState<boolean>(false);

    const theme = getTheme();

    const sharepointService: SharepointService = props.context.serviceScope.consume(SharepointService.serviceKey);
    const pdfService: PDFService = props.context.serviceScope.consume(PDFService.serviceKey);

    const handleFileNameChange = (e: React.FormEvent<HTMLInputElement>): void => {
        setFilename(e.currentTarget.value)
    }

    const handleCheck = (e: React.FormEvent<HTMLInputElement>): void => {
        if (e.currentTarget.name === "Delete") {
            setIsDeleteSelected(!isDeleteSelected)
        } else { setIsPreviewSelected(!isPreviewSelected) }
    }

    const handleDrag = (result: DropResult): void => {
        // pdfService.getAccessToken();
        const docs: IDocument[] = documents;
        const doc: IDocument = docs[result.source.index];
        if (result.destination && result.source.index !== result.destination.index) {
            docs.splice(result.source.index, 1);
            docs.splice(result.destination.index, 0, doc)
            setDocuments(docs);
        }
    }

    const handleMerge = async (): Promise<void> => {
        setIsLoading(true);
        try {
            const isFileExists = await sharepointService.checkFileExist(props.context.listView.folderInfo.folderPath, `${filename}.pdf`)
            if (isFileExists) {
                throw new Error(`Unable to merge: ${filename}.pdf already exists.`)
            }
            const fileRefs: string[] = await Promise.all(documents.map((document: IDocument) =>
                sharepointService.getFileRef(document.id)
            ));
            const fileContents: ArrayBuffer[] = await Promise.all(fileRefs.map((fileRef: string) =>
                sharepointService.getFileContent(fileRef)
            ));
            const mergedPDF: Uint8Array = await pdfService.mergePDFs(fileContents);
            const mergeResponse: JSON = await sharepointService.uploadFile(mergedPDF, filename, props.context.listView.folderInfo.folderPath);
            if (mergeResponse) {
                if (isDeleteSelected) {
                    await Promise.all(fileRefs.map((fileRef: string) => {
                        sharepointService.deleteFile(fileRef)
                    }));
                    setMessage(strings.MergeandDeleteInfo);
                }
                else {
                    setMessage(strings.MergeInfo);
                }
                if (isPreviewSelected) {
                    // window.open(`${props.context.pageContext.site.absoluteUrl.split('/sites')[0]}/${props.context.pageContext.list.serverRelativeUrl}/${filename}.pdf`, '_blank');
                    window.open(`${props.context.pageContext.site.absoluteUrl.split('/sites')[0]}/${mergeResponse}`, '_blank');
                }
            }
        }
        catch (error) {
            setMessage(error.message);
        }
        console.log(message)
        setIsLoading(false);
        setIsPanelOpen(false);
    }

    const renderFooter = (): React.ReactElement => {
        return (
            <div className={`${styles.flexbox} ${styles.justifyCenter} ${styles.footer}`}>
                <PrimaryButton onClick={() => handleMerge()}>{strings.Merge}</PrimaryButton>
                <DefaultButton onClick={props.closePanel}>{strings.Cancel}</DefaultButton>
            </div>
        )
    }

    return (
        <>
            <Panel isOpen={isPanelOpen} type={PanelType.medium} onDismiss={props.closePanel} headerText={strings.PanelHeader} onRenderFooterContent={() => renderFooter()} isFooterAtBottom={true}>
                {isLoading ? <div className={`${styles.flexbox} ${styles.alignCenter} ${styles.justifyCenter} ${styles.spinnerContainer}`}>
                    <Spinner label={strings.LoadingInfo} ariaLive="assertive" labelPosition="bottom" styles={{ label: { color: theme.palette.themePrimary }, circle: { borderTopColor: theme.palette.themePrimary, borderBottomColor: theme.palette.themeLight, borderLeftColor: theme.palette.themeLight, borderRightColor: theme.palette.themeLight } }} />
                </div> :
                    <div className={`${styles.flexbox} ${styles.panel}`}>
                        <DragDropContext onDragEnd={handleDrag}>
                            <Droppable droppableId='Documents'>
                                {
                                    (provided) => <div className={`${styles.docContainer} ${styles.flexbox} ${styles.alignCenter}`} ref={provided.innerRef} {...provided.droppableProps}>
                                        {documents.map((document: IDocument, index: number) =>
                                            <Document name={document.name} id={document.id} key={document.id} index={index} />
                                        )}
                                        {provided.placeholder}
                                    </div>
                                }
                            </Droppable>
                        </DragDropContext>
                        <div className={`${styles.flexbox} ${styles.pdfMergeInfo}`}>
                            <p>{strings.DragDropInfo}</p>
                            <p><strong>{strings.Note}</strong>{strings.PasswordProtectionInfo}</p>
                            <div className={`${styles.inputWrapper} ${styles.flexbox} ${styles.alignCenter}`}>
                                <span>{strings.NewFileName}</span>
                                <div className={`${styles.textboxWrapper} ${styles.flexbox} ${styles.alignCenter}`}>
                                    <input className={styles.textbox} type="text" value={filename} onChange={handleFileNameChange} />
                                    <span className={`${styles.pdfText} ${styles.flexbox} ${styles.alignCenter}`}>.pdf</span>
                                </div>
                            </div>
                            <div>
                                <div className={`${styles.checkboxWrapper} ${styles.flexbox} ${styles.alignCenter}`}>
                                    <input className={styles.checkbox} type='checkbox' id='Delete' name='Delete' defaultChecked={isDeleteSelected} onChange={handleCheck} />
                                    <label htmlFor='Delete'>{strings.DeleteInfo}</label>
                                </div>
                                <div className={`${styles.checkboxWrapper} ${styles.flexbox} ${styles.alignCenter}`}>
                                    <input className={styles.checkbox} type='checkbox' id='Preview' name='Preview' defaultChecked={isPreviewSelected} onChange={handleCheck} />
                                    <label htmlFor='Preview'>{strings.PreviewInfo}</label>
                                </div>
                            </div>
                        </div>
                    </div>}
            </Panel>
            <Dialog hidden={isPanelOpen} dialogContentProps={{ type: DialogType.close, subText: message }} modalProps={{ isBlocking: true, isDarkOverlay: false }} onDismiss={props.closePanel} />
        </>
    )
}

export default PDFMerge
import * as React from 'react';
import styles from './Document.module.scss';
import { Draggable } from 'react-beautiful-dnd';

export interface IDocumentProps {
    name: string,
    id:number
    index: number
}

const Document = (props: IDocumentProps):React.ReactElement => {

    return (
        <Draggable draggableId={props.id.toString()} index={props.index}>
            {
                (provided) => <div className={styles.document} ref={provided.innerRef} {...provided.draggableProps} {...provided.dragHandleProps}>{props.name}</div>
            }
        </Draggable>

    )
}

export default Document
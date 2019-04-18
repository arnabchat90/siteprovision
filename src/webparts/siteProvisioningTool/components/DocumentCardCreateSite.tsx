import * as React from 'react';
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
// import './DocumentCard.Example.scss';

const logo: any = require('../../../images/create-site.png');

export default class DocumentCardCreateSite extends React.Component<any, any> {
  public render(): JSX.Element {
    const previewProps: IDocumentCardPreviewProps = {
      previewImages: [
        {
          name: 'Create a new Team Site',
          previewImageSrc: logo,
          //iconSrc: logo,
          imageFit: ImageFit.cover,
          width: 350,
          height: 196
        }
      ]
    };

    return (
      <DocumentCard onClick={this.props.formLoadClick}>
        <DocumentCardPreview {...previewProps} />
        <DocumentCardTitle
          title="Create a new Team Site"
          shouldTruncate={true}
        />
        
      </DocumentCard>
    );
  }
}
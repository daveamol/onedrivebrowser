import * as React from 'react';
import * as MsGraph from '@microsoft/microsoft-graph-client';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import {
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  IColumn,
  ConstrainMode
} from 'office-ui-fabric-react/lib/DetailsList';
import {
  Breadcrumb, IBreadcrumbItem
} from 'office-ui-fabric-react/lib/Breadcrumb';
import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
import './DriveBrowser.css';

const fileIcons: string[] = [
  'accdb',
  'csv',
  'docx',
  'dotx',
  'mpp',
  'mpt',
  'odp',
  'ods',
  'odt',
  'one',
  'onepkg',
  'onetoc',
  'potx',
  'ppsx',
  'pptx',
  'pub',
  'vsdx',
  'vssx',
  'vstx',
  'xls',
  'xlsx',
  'xltx',
  'xsn'
];

export interface ListItem {
  id: string;
  fileName: string;
  extension: string;
  iconName: string;
  modifiedBy: string;
  dateModified: string;
  fileSize: string;
  data: any;
  type: listItemType;
}

interface DriveBrowserProperties {
  client: MsGraph.Client | null;
}

interface DriveBrowserState {
  listItems: ListItem[];
  BreadcrumbItems: IBreadcrumbItem[];
  isLoading: boolean;
}

enum listItemType {
  FILE,
  FOLDER
}

class DriveBrowser extends React.Component<DriveBrowserProperties, DriveBrowserState> {

  private columns: IColumn[];

  constructor(props: DriveBrowserProperties) {
    super(props);
    this.state = {
      listItems: [],
      BreadcrumbItems: [{ text: 'Home', key: 'root', isCurrentItem: true, onClick: this.onBreadCrumbSelected }],
      isLoading: true
    };

    this.columns = [
      {
        key: 'column1',
        name: 'File Type',
        headerClassName: 'DetailsListExample-header--FileIcon',
        className: 'DetailsListExample-cell--FileIcon',
        iconClassName: 'DetailsListExample-Header-FileTypeIcon',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: '',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item: ListItem) => {
          let data: any;
          if (item.type === listItemType.FOLDER || item.iconName.includes('ms-Icon')) {
            data = (
              <i className={item.iconName}>
                {name}
              </i>
            );
          } else {
            data = (
              <img
                src={item.iconName}
              />
            );
          }
          return (
            <span>
              {data}
            </span>
          );
        }
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: '',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
        onRender: (item: ListItem) => {
          let name: string = item.fileName;
          if (item.extension !== '') {
            name += '.' + item.extension;
          }
          let data;
          if (item.type === listItemType.FILE) {
            data = (
              <a href={item.data.webUrl} target="_blank" className="file-link">
                {name}
              </a>
            );
          } else {
            data = name;
          }
          return (
            <span>
              {data}
            </span>
          );
        }
      },
      {
        key: 'column3',
        name: 'Date Modified',
        fieldName: '',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        data: 'string',
        onRender: (item: ListItem) => {
          return (
            <span>
              {item.dateModified}
            </span>
          );
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Modified By',
        fieldName: '',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        isPadded: true,
        onRender: (item: ListItem) => {
          return (
            <span>
              {item.modifiedBy}
            </span>
          );
        },
      },
      {
        key: 'column5',
        name: 'File Size',
        fieldName: '',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onRender: (item: ListItem) => {
          return (
            <span>
              {item.fileSize}
            </span>
          );
        }
      },
    ];
  }

  componentDidUpdate(prevProps: DriveBrowserProperties, prevState: any) {
    if (this.props.client && prevProps.client !== this.props.client) {
      this.callGraphApi(this.state.BreadcrumbItems[this.state.BreadcrumbItems.length - 1].key);
    }
  }

  render() {
    let spinner;

    if (this.state.isLoading) {
      spinner = (
        <Spinner size={SpinnerSize.large} className="loading" />
      );
    } else {
      spinner = null;
    }

    return (
      <div>
        <Breadcrumb
          items={this.state.BreadcrumbItems}
        />
        <DetailsList
          items={this.state.listItems}
          columns={this.columns}
          onItemInvoked={this.onListItemInvoked}
          layoutMode={DetailsListLayoutMode.justified}
          checkboxVisibility={CheckboxVisibility.hidden}
          constrainMode={ConstrainMode.unconstrained}
          className="horizontal-scroll-hidden"
        />
        {spinner}
      </div>
    );
  }

  @autobind
  public onBreadCrumbSelected(ev: React.MouseEvent<HTMLElement>, breadcrumb: IBreadcrumbItem) {

    let pos: number = this.state.BreadcrumbItems.indexOf(breadcrumb);
    // First Check if use clicked on current active breadcrumb (which is the last breadcrumb)
    if (pos !== this.state.BreadcrumbItems.length - 1) {
      let slicedBreadcrumbs: IBreadcrumbItem[] = this.state.BreadcrumbItems.slice(0, pos + 1);
      slicedBreadcrumbs[pos].isCurrentItem = true;
      let urlPath = (pos === 0) ? breadcrumb.key : 'items/' + breadcrumb.key;
      this.callGraphApi(urlPath);
      this.setState({ BreadcrumbItems: slicedBreadcrumbs });
    }
  }

  private setItems(result: any) {
    let resultListItems: ListItem[] = [];

    for (let value of result.value) {
      let isFolder: boolean = false;

      if (value.folder) {
        isFolder = true;
      }

      let extension: string = isFolder ? '' : value.name.split('.')[1];
      let iconName: string = '';

      if (isFolder) {
        iconName = 'ms-Icon ms-Icon--FabricFolderFill icon-size';
      } else if (fileIcons.indexOf(extension) !== -1) {
        iconName =
          'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/'
          + extension
          + '_16x1.svg';
      } else {
        iconName = 'ms-Icon ms-Icon--Page icon-size';
      }

      resultListItems.push({
        id: value.id,
        fileName: isFolder ? value.name : value.name.split('.')[0],
        extension: extension,
        iconName: iconName,
        modifiedBy: value.lastModifiedBy.user.displayName,
        dateModified: (new Date(value.lastModifiedDateTime)).toString(),
        fileSize: isFolder ? '' : (Math.round((value.size / 1000) * 100 + Number.EPSILON) / 100).toString() + 'KB',
        data: value,
        type: isFolder ? listItemType.FOLDER : listItemType.FILE
      });
    }
    this.setState({ listItems: resultListItems, isLoading: false });
  }

  @autobind
  private onListItemInvoked(item: ListItem, index: number) {
    if (item.type === listItemType.FOLDER) {
      let urlPath: string = 'items/' + item.id;
      let breadcrumbItemsList: IBreadcrumbItem[] = this.state.BreadcrumbItems;

      this.callGraphApi(urlPath);
      breadcrumbItemsList.push(
        {
          text: item.fileName, key: item.id, isCurrentItem: true, onClick: this.onBreadCrumbSelected
        }
      );
      this.setState({ BreadcrumbItems: breadcrumbItemsList });
    }
  }

  @autobind
  private callGraphApi(urlPath: string) {
    if (this.props.client) {
      this.setState({ isLoading: true });
      this.props.client
        .api('/me/drive/' + urlPath + '/children')
        .get((err, res) => {
          console.log(res);
          this.setItems(res);
        });
    }
  }
}

export default DriveBrowser;
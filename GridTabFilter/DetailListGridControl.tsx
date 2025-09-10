import * as React from 'react';
import { IInputs } from "./generated/ManifestTypes";
import {
  Fabric,
  Link,
  Label,
  ScrollablePane,
  ScrollbarVisibility,
  ShimmeredDetailsList,
  Sticky,
  StickyPositionType,
  IRenderFunction,
  SelectionMode,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  ConstrainMode,
  IDetailsHeaderProps,
  TooltipHost,
  ITooltipHostProps,
  Stack
} from '@fluentui/react';
import { Tab, TabList } from "@fluentui/react-components";
import { initializeIcons } from '@fluentui/react/lib/Icons';
import lcid from 'lcid';

export interface IProps {
  pcfContext: ComponentFramework.Context<IInputs>,
  isModelApp: boolean,
  dataSetVersion: number
}

interface IColumnWidth {
  name: string,
  width: number
}

interface StatusTab {
  value: string;
  label: string;
}

initializeIcons();

export const DetailListGridControl: React.FC<IProps> = (props) => {  
  const tabField = props.pcfContext.parameters.TabOptionSetField?.raw || "statuscode";  
  const [columns, setColumns] = React.useState(getColumns(props.pcfContext));
  const [items, setItems] = React.useState(getItems(columns, props.pcfContext, tabField));
  const [isDataLoaded, setIsDataLoaded] = React.useState(props.isModelApp);
  const [selectedItemCount, setSelectedItemCount] = React.useState(0);  
  const [statuses, setStatuses] = React.useState<StatusTab[]>([]);
  const [selectedStatus, setSelectedStatus] = React.useState<string | null>(null);

  // Set the isDataLoaded state based upon the paging totalRecordCount
  React.useEffect(() => {
    const dataSet = props.pcfContext.parameters.sampleDataSet;
    if (dataSet.loading || props.isModelApp) return;
    setIsDataLoaded(dataSet.paging.totalResultCount !== -1);
  }, [items, props.isModelApp, props.pcfContext]);

  // When the component is updated this will determine if the sampleDataSet has changed.
  React.useEffect(() => {
    setItems(getItems(columns, props.pcfContext, tabField));
  }, [props.dataSetVersion, columns, props.pcfContext, tabField]);

  // When the component is updated this will determine if the width of the control has changed.
  React.useEffect(() => {
    setColumns(updateColumnWidths(columns, props.pcfContext));
  }, [props.pcfContext.mode.allocatedWidth]);

  React.useEffect(() => {
    const statusSet = new Map<string, string>();
    props.pcfContext.parameters.sampleDataSet.sortedRecordIds.forEach(key => {
      const record = props.pcfContext.parameters.sampleDataSet.records[key];
      const value = record.getValue(tabField);
      const label = record.getFormattedValue(tabField);
      if (value !== undefined && !statusSet.has(String(value))) {
        statusSet.set(String(value), label);
      }
    });
    const statusArr = Array.from(statusSet.entries()).map(([value, label]) => ({ value, label }));
    setStatuses(statusArr);
    if (statusArr.length > 0 && (selectedStatus === null || !statusSet.has(selectedStatus))) {
      setSelectedStatus(statusArr[0].value);
    }
  }, [items, tabField]);

  // the selector used by the DetailList
  const _selection = React.useMemo(() =>
    new Selection({
      onSelectionChanged: () => {
        _setSelectedItemsOnDataSet();
      }
    }), [props.pcfContext.parameters.sampleDataSet]
  );

  const _setSelectedItemsOnDataSet = () => {
    const selectedKeys = [];
    const selections = _selection.getSelection();
    for (const selection of selections) {
      selectedKeys.push(selection.key as string);
    }
    setSelectedItemCount(selectedKeys.length);
    props.pcfContext.parameters.sampleDataSet.setSelectedRecordIds(selectedKeys);
  }
  
  const _onColumnClick = (ev?: React.MouseEvent<HTMLElement>, column?: IColumn): void => {
    let isSortedDescending = column?.isSortedDescending;

    if (column?.isSorted) {
      isSortedDescending = !isSortedDescending;
    }

    setItems(copyAndSort(items, column?.fieldName!, props.pcfContext, isSortedDescending));
    setColumns(
      columns.map(col => {
        col.isSorted = col.key === column?.key;
        col.isSortedDescending = isSortedDescending;
        return col;
      })
    );
  }

  const _onRenderDetailsHeader = (props: IDetailsHeaderProps | undefined, defaultRender?: IRenderFunction<IDetailsHeaderProps>): JSX.Element => {
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
        {defaultRender!({
          ...props!,
          onRenderColumnHeaderTooltip: (tooltipHostProps: ITooltipHostProps | undefined) => <TooltipHost {...tooltipHostProps} />
        })}
      </Sticky>
    )
  }

  // Filter grid items based on selected tab/status
  const filteredItems = React.useMemo(() => {
    if (!selectedStatus) return items;
    return items.filter(item => String(item[`${tabField}_raw`]) === selectedStatus);
  }, [items, selectedStatus, tabField]);

  return (
    <Stack grow
      styles={{
        root: {
          width: "100%",
          height: "inherit",
        },
      }}>
      <Stack.Item>
        <div style={{ padding: "8px 0" }}>
          <TabList
            selectedValue={selectedStatus}
            onTabSelect={(_, data) => setSelectedStatus(data.value as string)}
          >
            {statuses.map(status => (
              <Tab key={status.value} value={status.value}>{status.label}</Tab>
            ))}
          </TabList>
        </div>
      </Stack.Item>
      <Stack.Item
        verticalFill
        styles={{
          root: {
            height: "100%",
            overflowY: "auto",
            overflowX: "auto",
          },
        }}
      >
        <div
          style={{ position: 'relative', height: '100%' }}>
          <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
            <ShimmeredDetailsList
              enableShimmer={!isDataLoaded}
              className='list'
              items={filteredItems}
              columns={columns}
              setKey="set"
              selection={_selection}
              onColumnHeaderClick={_onColumnClick}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="Row checkbox"
              selectionMode={SelectionMode.multiple}
              onRenderDetailsHeader={_onRenderDetailsHeader}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={ConstrainMode.unconstrained}
            />
          </ScrollablePane>
        </div>
      </Stack.Item>
      <Stack.Item align="start">
        <div className="detailList-footer">
          <Label className="detailList-gridLabels">Records: {filteredItems.length.toString()} ({selectedItemCount} selected)</Label>
        </div>
      </Stack.Item>
    </Stack>
  );
};

const navigate = (item: any, linkReference: string | undefined, pcfContext: ComponentFramework.Context<IInputs>) => {
  pcfContext.parameters.sampleDataSet.openDatasetItem(item[linkReference + "_ref"])
};

// get the items from the dataset
const getItems = (columns: IColumn[], pcfContext: ComponentFramework.Context<IInputs>, tabField: string) => {
  const dataSet = pcfContext.parameters.sampleDataSet;

  const resultSet = dataSet.sortedRecordIds.map(function (key) {
    const record = dataSet.records[key];
    const newRecord: any = {
      key: record.getRecordId()
    };

    for (const column of columns) {
      newRecord[column.key] = record.getFormattedValue(column.key);
      if (isEntityReference(record.getValue(column.key))) {
        const ref = record.getValue(column.key) as ComponentFramework.EntityReference;
        newRecord[column.key + '_ref'] = ref;
      }
      else if (column.data.isPrimary) {
        newRecord[column.key + '_ref'] = record.getNamedReference();
      }
    }    
    newRecord[`${tabField}_raw`] = record.getValue(tabField);
    return newRecord;
  });

  return resultSet;
}

// get the columns from the dataset
const getColumns = (pcfContext: ComponentFramework.Context<IInputs>): IColumn[] => {
  const dataSet = pcfContext.parameters.sampleDataSet;
  const iColumns: IColumn[] = [];

  const columnWidthDistribution = getColumnWidthDistribution(pcfContext);

  for (const column of dataSet.columns) {
    const iColumn: IColumn = {
      key: column.name,
      name: column.displayName,
      fieldName: column.alias,
      currentWidth: column.visualSizeFactor,
      minWidth: 5,
      maxWidth: columnWidthDistribution.find(x => x.name === column.alias)?.width || column.visualSizeFactor,
      isResizable: true,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      className: 'detailList-cell',
      headerClassName: 'detailList-gridLabels',
      data: { isPrimary: column.isPrimary }
    };

    // Show formatted value for option set columns (including any tab field)
    if (
      column.dataType === "OptionSet" ||
      column.dataType === "State" ||
      column.dataType === "Status"
    ) {
      iColumn.onRender = (item: any, index: number | undefined, col?: IColumn) => (
        <span>{item[col!.fieldName!]}</span>
      );
    }
    //create links for primary field and entity reference.
    else if (column.dataType.startsWith('Lookup.') || column.isPrimary) {
      iColumn.onRender = (item: any, index: number | undefined, column: IColumn | undefined) => (
        <Link key={item.key} onClick={() => navigate(item, column!.fieldName, pcfContext)}>{item[column!.fieldName!]}</Link>
      );
    }
    else if (column.dataType === 'SingleLine.Email') {
      iColumn.onRender = (item: any, index: number | undefined, column: IColumn | undefined) => (
        <Link href={`mailto:${item[column!.fieldName!]}`}>{item[column!.fieldName!]}</Link>
      );
    }
    else if (column.dataType === 'SingleLine.Phone') {
      iColumn.onRender = (item: any, index: number | undefined, column: IColumn | undefined) => (
        <Link href={`skype:${item[column!.fieldName!]}?call`}>{item[column!.fieldName!]}</Link>
      );
    }
    
    const isSorted = dataSet?.sorting?.findIndex(s => s.name === column.name) !== -1 || false
    iColumn.isSorted = isSorted;
    if (isSorted) {
      iColumn.isSortedDescending = dataSet?.sorting?.find(s => s.name === column.name)?.sortDirection === 1 || false;
    }

    iColumns.push(iColumn);
  }
  return iColumns;
}

const getColumnWidthDistribution = (pcfContext: ComponentFramework.Context<IInputs>): IColumnWidth[] => {
  const widthDistribution: IColumnWidth[] = [];
  const columnsOnView = pcfContext.parameters.sampleDataSet.columns;
  const totalWidth: number = pcfContext.mode.allocatedWidth - 250;
  let widthSum = 0;

  columnsOnView.forEach(function (columnItem) {
    widthSum += columnItem.visualSizeFactor;
  });

  let remainWidth: number = totalWidth;

  columnsOnView.forEach(function (item, index) {
    let widthPerCell = 0;
    if (index !== columnsOnView.length - 1) {
      const cellWidth = Math.round((item.visualSizeFactor / widthSum) * totalWidth);
      remainWidth = remainWidth - cellWidth;
      widthPerCell = cellWidth;
    }
    else {
      widthPerCell = remainWidth;
    }
    widthDistribution.push({ name: item.alias, width: widthPerCell });
  });

  return widthDistribution;
}

const updateColumnWidths = (columns: IColumn[], pcfContext: ComponentFramework.Context<IInputs>): IColumn[] => {
  const columnWidthDistribution = getColumnWidthDistribution(pcfContext);
  const currentColumns = columns;

  return currentColumns.map(col => {
    const newMaxWidth = columnWidthDistribution.find(x => x.name === col.fieldName);
    if (newMaxWidth) col.maxWidth = newMaxWidth.width;
    return col;
  });
}

//sort the items in the grid.
const copyAndSort = <T,>(items: T[], columnKey: string, pcfContext: ComponentFramework.Context<IInputs>, isSortedDescending?: boolean): T[] => {
  const key = columnKey as keyof T;
  const sortedItems = items.slice(0);
  sortedItems.sort((a: T, b: T) => (a[key] || '' as any).toString().localeCompare((b[key] || '' as any).toString(), getUserLanguage(pcfContext), { numeric: true }));

  if (isSortedDescending) {
    sortedItems.reverse();
  }

  return sortedItems;
}

const getUserLanguage = (pcfContext: ComponentFramework.Context<IInputs>): string => {
  const language = lcid.from(pcfContext.userSettings.languageId);
  return language ? language.substring(0, language.indexOf('_')) : 'en';
}

const isEntityReference = (obj: any): obj is ComponentFramework.EntityReference => {
  return typeof obj?.etn === 'string';
}
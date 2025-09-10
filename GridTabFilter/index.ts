/* eslint-disable no-prototype-builtins */
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import * as React from 'react';
import * as ReactDOM from 'react-dom/client';
import {IProps, DetailListGridControl}  from './DetailListGridControl'

export class GridTabFilter implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _detailList: HTMLDivElement;
	private _dataSetVersion: number;
	private _isModelApp: boolean
	private _props: IProps;
    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        context.mode.trackContainerResize(true);
		this._container = container;
		this._context = context;	
		this._isModelApp = window.hasOwnProperty('getGlobalContextObject');
		this._dataSetVersion = 0;

		this._props = {
			pcfContext: this._context,
			isModelApp: this._isModelApp,
			dataSetVersion: this._dataSetVersion
		}
		
		this._container.style.position = 'relative';
		this._detailList = document.createElement("div");
		this._detailList.setAttribute("id", "detailList");		
		this._detailList.setAttribute("data-is-scrollable", "true");
		
		if (this._context.mode.allocatedHeight !== -1)
		{
			this._detailList.style.height = `${(this._context.mode.allocatedHeight).toString()}px`
		}
		else
		{						
			const rowspan = (this._context.mode as any).rowSpan;
			if (rowspan) this._detailList.style.height = `${(rowspan * 1.5).toString()}em`;
		}

		this._container.appendChild(this._detailList);		
		context.parameters.sampleDataSet.paging.setPageSize(5000);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const dataSet = context.parameters.sampleDataSet;		
		if (dataSet.loading) return;		
		if (!this._isModelApp)
		{
			//since we are in a canvas app let's make sure we set the height of the control
			this._detailList.style.height = `${(this._context.mode.allocatedHeight).toString()}px`						
			dataSet.paging.setPageSize(dataSet.paging.totalResultCount);
		}
		
		if (this._isModelApp && dataSet.paging.hasNextPage) {
			dataSet.paging.loadNextPage();
			return;
		}
		
		this._props.dataSetVersion = this._dataSetVersion++;				
		if (!(this._detailList as any)._root) {
			(this._detailList as any)._root = ReactDOM.createRoot(this._detailList);
		}
		(this._detailList as any)._root.render(
			React.createElement(DetailListGridControl, this._props)
		);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {};
    }
   
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}

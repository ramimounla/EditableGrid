import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
// import * as $ from 'jquery';

export class EditableGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _container: HTMLDivElement;

	/**
	 * Empty constructor.
	 */
	constructor() {
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		// Add control initialization code
		this._container = document.createElement("div");
		this._container.className = "table-like";

		container.appendChild(this._container);
	}


	private sanitizeNameToCss(name: string): string {
		return name.toLowerCase().replace(' ', '-');
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view

		if (!context.parameters.recordSet.loading) {

			this._container.innerHTML = "";
			var recordSet = context.parameters.recordSet;

			var headers = <HTMLDivElement>document.createElement("div");
			headers.className = "header";
			context.parameters.recordSet.columns.forEach(column => {
				var span = <HTMLSpanElement>document.createElement("span");
				column.displayName !== 'Tags' && (span.className = "element "); //if not equals tag
				span.innerText = column.displayName;
				headers.appendChild(span);
			});
			this._container.appendChild(headers);


			recordSet.sortedRecordIds.forEach(recordId => {
				var recordDiv = <HTMLDivElement>document.createElement("div");
				recordDiv.className = "row";
				context.parameters.recordSet.columns.forEach(column => {
					
					recordDiv.id = recordId;
					var span = <HTMLSpanElement>document.createElement("span");
					span.className = "element " + this.sanitizeNameToCss(column.displayName);

					var input = <HTMLInputElement>document.createElement("input");

					if (column.dataType === "Lookup.Simple" && column.name === "dxc_variation") { 

						

						//@ts-ignore
						input.value = recordSet.records[recordId].getValue(column.name) === null ? "" : recordSet.records[recordId].getValue(column.name).name;

						input.addEventListener('click', e => {

							//@ts-ignore
							var jurisdiction = Xrm.Page.getAttribute("dxc_juridiction").getValue()[0].id;
							//@ts-ignore
							var recordId = input.parentElement.parentElement.id;
							//@ts-ignore
							var gameId = context.parameters.recordSet.records[recordId].getValue("dxc_game").id.guid;

							var filter  = "<filter type='and'><condition attribute='statecode' operator='eq' value='0' /><condition attribute='dxc_jurisdiction' operator='eq' value='" + jurisdiction + "' /><condition attribute='dxc_game' operator='eq' value='" + gameId + "' /></filter>";

							//@ts-ignore
							var lookupOptions = {
								defaultEntityType: "dxc_variationjurisdiction",
								entityTypes: ["dxc_variationjurisdiction"],
								disableMru: true,
								allowMultiSelect: false,

								filters: [{
									filterXml: filter,
									entityLogicalName: "dxc_variationjurisdiction"
								}]
							};

							//@ts-ignore
							Xrm.Utility.lookupObjects(lookupOptions)
								.then(function (result: any) {
									if (result.length > 0) {
										input.value = result[0].name;
									}
								})
								.fail(function (error: any) {
									alert(error);
								});

						});


					}
					else if (column.dataType === "Lookup.Simple") {
						//@ts-ignore
						input.value = recordSet.records[recordId].getValue(column.name) === null ? "" : recordSet.records[recordId].getValue(column.name).name;

						input.addEventListener('click', e => {
							//@ts-ignore
							var lookupOptions = {
								defaultEntityType: column.name,
								entityTypes: [column.name],
								allowMultiSelect: false,
							};

							//@ts-ignore
							Xrm.Utility.lookupObjects(lookupOptions)
								.then(function (result: any) {
									if (result.length > 0) {
										input.value = result[0].name;
									}
								})
								.fail(function (error: any) {
									alert(error);
								});

						});
					}
					else {
						input.value = <string>recordSet.records[recordId].getValue(column.name);

					}



					span.appendChild(input);
					recordDiv.appendChild(span);

				});
				this._container.appendChild(recordDiv);
			});

		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}
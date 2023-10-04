namespace Helpers.Data {
    /**
     * Retrieve a record from the web api
     * @param entityName The name of the entity without the trailing 's'
     * @param id Guid of the record
     * @param cols columns to retrieve
     * @param keyAttribute keyAttribute defaults to entity + 'id', use this to override key attribute (ex activity)
     */
    export function Retrieve(entityName: string, id: string, cols?: Array<string>, keyAttribute?: string) {
        if (!keyAttribute) {
            keyAttribute = entityName + "id";
        }

        if (entityName.substr(entityName.length - 1) == "y"
            && !entityName.endsWith('journey'))
            entityName = entityName.substr(0, entityName.length -1) + "ies";
        else
            entityName = entityName + "s";

        id = id.cleanGuid();

        var select = "$select=" + keyAttribute;

        if (cols) {
            select += ",";
            select += cols.join(',');
        }

        var req = new XMLHttpRequest();
        var clientURL = Helpers.GlobalHelper.GetContext().getClientUrl();
        req.open("GET", encodeURI(clientURL + "/api/data/v9.2/" + entityName + "(" + id + ")?" + select), false);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"');
        req.send(null);
        return JSON.parse(req.responseText);
    }

    export function RetrieveWithCustomFilter(urlEnding: string) {
        var req = new XMLHttpRequest();
        var clientURL = Helpers.GlobalHelper.GetContext().getClientUrl();
        req.open("GET", clientURL + "/api/data/v9.2/" + urlEnding, false);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"');
        req.send(null);
        return JSON.parse(req.responseText);
    }

    /**
     * Retrieve a record from the web api
     * @param entityName The name of the entity without the trailing 's'
     * @param id Guid of the record
     * @param cols columns to retrieve
     * @param keyAttribute keyAttribute defaults to entity + 'id', use this to override key attribute (ex activity)
     */
    export async function RetrieveAsync(entityName: string, id: string, cols?: Array<string>, keyAttribute?: string) {
        if (!keyAttribute) {
            keyAttribute = entityName + "id";
        }
        if (entityName.substr(entityName.length - 1) == "y"
            && !entityName.endsWith('journey'))
            entityName = entityName.substr(0,entityName.length-1) + "ies";
        else
            entityName = entityName + "s";

        id = id.cleanGuid();

        var records = await Data.RetrieveMultipleAsync(entityName + "?$filter=" + keyAttribute + " eq " + id, true, cols);
        if (records.length == 0) {
            throw new Error("RetrieveAsync error: Entity not found.");
        }
        if (records.length > 1) {
            throw new Error("RetrieveAsync error: Found more then 1");
        }

        return records[0];
    }

    export function RetrieveWithCustomFilterAsync(urlEnding: string) {
        return new Promise((resolve: (records: Array<any>) => any, reject: (statusText: string) => any) => {
            var req = new XMLHttpRequest();
            req.open("GET", Helpers.GlobalHelper.GetContext().getClientUrl() + "/api/data/v9.2/" + urlEnding, true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"");

            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var results = JSON.parse(this.response);
                        resolve(results);
                    }
                    else {
                        reject(this.statusText);
                    }
                }
            };
            req.send();
        });
    }

    export function RetrieveMultipleAsync(urlEnding: string, addFormattedValue: boolean = true, cols?: Array<string>) {
        return new Promise((resolve: (records: Array<any>) => any, reject: (statusText: string) => any) => {
            if (cols) {
                if (urlEnding.indexOf('?') == -1) {
                    urlEnding += "?";
                }
                else {
                    urlEnding += "&";
                }

                urlEnding += "$select=" + cols.join(',');
            }

            var req = new XMLHttpRequest();
            req.open("GET", Helpers.GlobalHelper.GetContext().getClientUrl() + "/api/data/v9.2/" + urlEnding, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            if (addFormattedValue) {
                //req.setRequestHeader("Prefer", "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"");
                req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            }
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var results = JSON.parse(this.response).value;
                        resolve(results);
                    }
                    else {
                        reject(this.statusText);
                    }
                }
            };

            req.send();
        });
    }

	/**
	 * Get data from the OData services using fetch XML. Try to use RetrieveFetchAsync as much as possible
	 * @param entityName name of the enntity
	 * @param xml fetch XML
	 */
	export function RetrieveFetch(entityName: string, xml: string) : any[] {

        if (entityName.substr(entityName.length - 1) == "y"
            && !entityName.endsWith('journey'))
            entityName = entityName.substr(0, -1) + "ies";
        else
		    entityName = entityName + "s";
		
		var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.2/" + entityName + "?fetchXml=" + encodeURIComponent(xml), false);
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
		req.send(null);

        return JSON.parse(req.responseText).value;
    }


	/**
	 * Get data from the OData services using fetch XML
	 * @param entityName name of the enntity
	 * @param xml fetch XML
	 */
	export function RetrieveFetchAsync(entityName: string, xml: string): Promise<any[]> {
        return new Promise((resolve: (records: Array<any>) => any, reject: (statusText: string) => any) => {
            if (entityName.substr(entityName.length - 1) == "y"
                && !entityName.endsWith('journey'))
                entityName = entityName.substr(0, entityName.length -1) + "ies";
            else
			    entityName = entityName + "s";
			var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.2/" + entityName + "?fetchXml=" + encodeURIComponent(xml), true);
			req.setRequestHeader("OData-MaxVersion", "4.0");
			req.setRequestHeader("OData-Version", "4.0");
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
			req.onreadystatechange = function () {
				if (this.readyState === 4) {
					req.onreadystatechange = null;
					if (this.status === 200) {
						var results = JSON.parse(this.response);
						resolve(results.value);
					} else {
						reject(this.statusText);
					}
				}
			};
			req.send();
		});
	}

    export function UpdateAsync(urlEnding: string, entity: {}) {
        return new Promise((resolve: (succes: Boolean) => any, reject: (statusText: string) => any) => {
            var req = new XMLHttpRequest();
            req.open("PATCH", Helpers.GlobalHelper.GetContext().getClientUrl() + "/api/data/v9.2/" + urlEnding, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 204) {
                        //Success - No Return Data - Do Something
                        resolve(true);
                    }
                    else {
                        reject(this.response);
                    }
                }
            };
            req.send(JSON.stringify(entity));
        });
    }

    export function CreateAsync(entityName: string, entity: {}) {
        return new Promise((resolve: (succes: string) => any, reject: (statusText: string) => any) => {
            entityName = entityName + "s";

            var req = new XMLHttpRequest();
            req.open("POST", Helpers.GlobalHelper.GetContext().getClientUrl() + "/api/data/v9.2/" + entityName, true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 204) {
                        //Success - Return Data - EntityId
                        var result = this.getResponseHeader("OData-EntityId")
                        var entityId = result.split(entityName)[1].replace(/[()]/g, "").toLowerCase();
                        resolve(entityId);
                    }
                    else {
                        reject(JSON.parse(this.response));
                    }
                }
            };
            req.send(JSON.stringify(entity));
        });
    }

    export function AssociateRecord(urlEnding: string, association: Object) {
        return new Promise((resolve: (succes: boolean) => any, reject: (statusText: string) => any) => {
            var req = new XMLHttpRequest();
            req.open("POST", Helpers.GlobalHelper.GetContext().getClientUrl() + "/api/data/v9.2/" + urlEnding, true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 204) {
                        //Success
                        resolve(true);
                    }
                    else {
                        reject(this.statusText);
                    }
                }
            };
            req.send(JSON.stringify(association));
        });
    }

    export function DisassociateRecord(urlEnding: string) {
        return new Promise((resolve: (succes: boolean) => any, reject: (statusText: string) => any) => {
            var req = new XMLHttpRequest();
            req.open("DELETE", Helpers.GlobalHelper.GetContext().getClientUrl() + "/api/data/v9.2/" + urlEnding, true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 204) {
                        //Success
                        resolve(true);
                    }
                    else {
                        reject(this.statusText);
                    }
                }
            };
            req.send();
        });
    }

    export function ExecuteGlobalAction<T>(customActionName: string, parameters: {}): Promise<T> {
        return new Promise((resolve: (result: T) => any, reject: (statusText: string) => any) => {
            var req = new XMLHttpRequest();
            req.open("POST", Helpers.GlobalHelper.GetContext().getClientUrl() + "/api/data/v9.2/" + customActionName, true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var output = JSON.parse(this.response);
                        resolve(output as T);
                    }
                    else {
                        reject(this.statusText);
                    }
                }
            };
            req.send(JSON.stringify(parameters));
        });
    }

    export function ExecuteAction<T>(entityName: string, id: string, actionName: string, parameters: {}): Promise<T> {

        return new Promise((resolve: (result: T) => any, reject: (statusText: string) => any) => {

            var req = new XMLHttpRequest();
            req.open("POST", `${Helpers.GlobalHelper.GetContext().getClientUrl()}/api/data/v9.2/${entityName}s(${id.cleanGuid()})/Microsoft.Dynamics.CRM.${actionName}`, true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var output = JSON.parse(this.response);
                        resolve(output as T);
                    }
                    else {
                        reject(this.statusText);
                    }
                }
            };
            req.send(JSON.stringify(parameters));
        });
    }
}

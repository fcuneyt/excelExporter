var excelExporter = {};
(function () {
	var allKeys = [];
	var keyMap = [];
	var type = "data:application/vnd.ms-excel;charset=utf-8,";
	var template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>';
	var getMapIndex = function (key) {
		var index = -1;
		for (var j = 0; j < keyMap.length; j++) {
			var currentKey = keyMap[j];
			if (currentKey[key]) {
				index = keyMap.indexOf(currentKey);
				break;
			}
		}
		return index;
	};
	var pushKeys = function (keyList, parentObject) {
		for (var i = 0; i < keyList.length; i++) {
			var keyObject = {
				key: keyList[i],
				parent: parentObject
			};
			var isExist = false;
			for (var j = 0; j < allKeys.length; j++) {
				var currentKey = allKeys[j];
				var keyExist = currentKey.key === keyObject.key
					&& currentKey.parent === keyObject.parent;
				var mapIndex = getMapIndex((keyObject.parent.length > 0
						? keyObject.parent + "."
						: "")
					+ keyObject.key);
				if (keyExist || (keyMap.length > 0 && mapIndex === -1)) {
					isExist = true;
					break;
				}
			}
			if (!isExist) {
				allKeys.push(keyObject);
			}
		}
	};
	var getKeyIndex = function (key, parent) {
		var index = -1;
		for (var j = 0; j < allKeys.length; j++) {
			var currentKey = allKeys[j];
			if (currentKey.key === key
				&& currentKey.parent === parent) {
				index = allKeys.indexOf(currentKey);
				break;
			}
		}
		return index;
	}
	var removeKey = function (key, parent) {
		var index = getKeyIndex(key, parent);
		if (index !== -1) {
			allKeys.splice(index, 1);
		}
	};
	var rowList = [];
	var fillDataRow = function (row, dataItem, keys, parent) {
		var currentRowList = [];
		for (var i = 0; i < keys.length; i++) {
			var currentKey = keys[i];
			var dataKey = dataItem[currentKey];
			var mapIndex = getMapIndex((parent.length > 0
					? parent + "."
					: "")
				+ currentKey);
			if (Object.prototype.toString.call(dataKey) === "[object Array]") {
				var childKeys = Object.keys(dataKey[0]);
				removeKey(currentKey, parent);
				pushKeys(childKeys, currentKey);
				for (var j = 0; j < dataKey.length; j++) {
					if (keyMap.length > 0) {
						var childMapIndex = -1;
						for (var l = 0; l < childKeys.length; l++) {
							childMapIndex = getMapIndex(currentKey + "." + childKeys[l]);
							if (childMapIndex !== -1) {
								break;
							}
						}
						if (childMapIndex !== -1) {
							var childRow = document.createElement("tr");
							//calculate indent and insert empty cells.
							for (var k = 0; k < childMapIndex; k++) {
								childRow.appendChild(document.createElement("td"));
							}
							currentRowList = fillDataRow(childRow, dataKey[j], childKeys, currentKey);
							currentRowList.push(childRow);
						}
					} else {
						var childRow = document.createElement("tr");
						var keyIndex = getKeyIndex(childKeys[0], currentKey);
						//calculate indent and insert empty cells.
						for (var m = 0; m < keyIndex; m++) {
							childRow.appendChild(document.createElement("td"));
						}
						currentRowList = fillDataRow(childRow, dataKey[j], childKeys, currentKey);
						currentRowList.push(childRow);
					}
				}
			} else if (typeof (dataKey) === "object") {
				var childKeys = Object.keys(dataKey);
				removeKey(currentKey, parent);
				pushKeys(childKeys, currentKey);
				currentRowList = fillDataRow(row, dataKey, childKeys, currentKey);
			} else {
				if (keyMap.length === 0 || mapIndex !== -1) {
					var cell = document.createElement("td");
					var value = dataKey;
					cell.innerHTML = value
						? value
						: "";
					row.appendChild(cell);
				}
			}
		}
		rowList.concat(currentRowList);
		return rowList;
	};
	var getDataRowList = function (dataList, keys, parent) {
		var dataRowList = [];
		for (var i = 0; i < dataList.length; i++) {
			rowList = [];
			var row = document.createElement("tr");
			var children = fillDataRow(row, dataList[i], keys, parent);
			dataRowList.push(row);
			for (var j = 0; j < children.length; j++) {
				dataRowList.push(children[j]);
			}
		}
		return dataRowList;
	};
	var fillTable = function (dataList, keys, parent) {
		var table = document.createElement("table");
		table.id = "Hede";

		var dataRowList = getDataRowList(dataList, keys, parent);
		for (var i = 0; i < dataRowList.length; i++) {
			table.appendChild(dataRowList[i]);
		}

		var headerRow = table.insertRow(0);
		for (var j = 0; j < allKeys.length; j++) {
			var cell = document.createElement("td");
			var keyObject = allKeys[j];
			var mapKey = (keyObject.parent.length > 0
					? keyObject.parent + "."
					: "")
				+ keyObject.key;
			var mapInfo = keyMap[getMapIndex(mapKey)];
			if (mapInfo) {
				cell.innerText = mapInfo[mapKey];
			} else {
				cell.innerText = mapKey;
			}
			headerRow.appendChild(cell);
		}

		return table;
	};
	var stringify = function (table) {
		return encodeURIComponent(template.replace("{table}", table.innerHTML));
	};
	var downloadFile = function (table, fileName) {
		if (!fileName) {
			var dt = new Date();
			var yyyy = dt.getFullYear().toString();
			var mm = (dt.getMonth() + 1).toString();
			var dd = dt.getDate().toString();
			var hh = dt.getHours().toString();
			var min = dt.getMinutes().toString();
			fileName = "export_" + yyyy + (mm[1] ? mm : "0" + mm[0]) + (dd[1] ? dd : "0" + dd[0]) + "-" + (hh[1] ? hh : "0" + hh[0]) + (min[1] ? min : "0" + min[0]) + ".xls";
		}
		if (fileName.indexOf(".xls") === -1) {
			filename += ".xls";
		}
		var exportData = stringify(table);
		var downloadLink = document.createElement("a");
		downloadLink.href = type + exportData;
		downloadLink.download = fileName;
		document.body.appendChild(downloadLink);
		downloadLink.click();
		document.body.removeChild(downloadLink);
	};
	this.fromTable = function (tableId, fileName) {
		var table = document.getElementById(tableId);
		downloadFile(table, fileName);
	};
	this.fromJson = function (data, selector, dataMap, fileName) {
		if (dataMap && dataMap.length > 0) {
			keyMap = dataMap;
		}
		if (typeof (data) !== "object") {
			try {
				data = JSON.parse(data);
			} catch (e) {
				throw e;
			}
		}
		allKeys = [];
		var dataObject = data[selector];

		var keys;
		var exportData;
		if (Object.prototype.toString.call(dataObject) === "[object Array]") {
			keys = Object.keys(dataObject[0]);
			exportData = dataObject;
		} else {
			keys = Object.keys(dataObject);
			exportData = [dataObject];
		}

		pushKeys(keys, "");
		downloadFile(fillTable(exportData, keys, ""), fileName);
	};
}).apply((excelExporter));

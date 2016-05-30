# excelExporter
A pure js excel exporter.

This library can export your data (from html table or JSON data) to excel.

#Usage

##With JSON Data
```javascript
$.getJSON('Your-JSON-URL', null, function (response) {
  //Give it your data. Define which data to export e.g. 'Items'.
	excelExporter.fromJson(response, 'Items',
	//And give it your mapping if you want.
	[
		{ "CityId": "City No" },
		{ "Description": "City Name" },
		{ "CountryReference.Description": "Country Name" },
		{ "TownList.Description": "Town Names"}
	]);
});
```
Result with sampleData.Json you will get an excel like as follows.

City No |	City Name	|	Country Name	|	Town Names
-------	| -------------		|	------------	|	-----
1	|	les Escaldes	|	Europe/Andorra      
	|			|			| First Town
	|			|			| Second Town
2	|	Andorra la Vella	|	Europe/Andorra
	|			|			| Third Town
3	|	Umm al Qaywayn	|	Asia/Dubai
4	|	Ras al-Khaimah	|	Asia/Dubai
5	|	Khawr Fakkān	|	Asia/Dubai

##With Table Selector
```html
<table id="data-table">
	<thead>
		<tr>
			<td>Data</td>
			<td>Second Data</td>
		</tr>
	</thead>
	<tr><td>a1</td></tr>
	<tr>
		<td></td>
		<td>
			<table>
				<tr>
					<td>b1</td>
				</tr>
				<tr>
					<td>b2</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr><td>a2</td></tr>
	<tr>
		<td></td>
		<td>
			<table>
				<tr>
					<td>c1</td>
				</tr>
				<tr>
					<td>c2</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
```
	
```javascript
excelExporter.fromTable('data-table');
```

Result is.

Data		|	Second Data
------------	|	-------------
a1		|
		|	b1
		|	b2
a2		|
		|	c1
		|	c2

#Installation

##Nuget
```
Install-Package excelExporter.js
```

##Bower
```
bower install excelexporter
```

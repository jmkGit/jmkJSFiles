(function () {
	// Note this seems to get called a lot, no idea why as I don't understand enough about Qlikview and how that refreshes its controls
	const styles = `color: green; font-size: 2em;`;
	console.log('%c** SSO Fix version 11/08/22 08:37 **', styles);
	var objs = [];
	var aspNetEndpoint = "http://edpqlvap01.millerextra.com/ssoQlikview/GenerateSso.ashx";
	//Live = "https://qlikview.millerextra.com/ssoQlikview/GenerateSso.ashx";

	function tableToString(table) {
		console.log("getting table", table);
		if (table.Type == "TX") {
			// handle single text box type (current week)
			return "|1|1|" + table.GetText() + "|";
		}
		var s = [];
		if (table.Layout.Caption != undefined && table.Layout.Caption.label != undefined) {
			s.push(table.Layout.Caption.label);
		} else {
			s.push('UNNAMED'); // push empty string or the structure is lost
		}

		//row count
		var rc = table.Data.Rows.length + table.Data.HeaderRows.length;
		s.push(rc);
		//col count
		if (rc > 0) {
			s.push(table.Data.Rows[0].length);
		} else {
			s.push(0);
		}
		//col headers
		if (table.Data.HeaderRows.length > 0) {
			for (var col = 0; col < table.Data.HeaderRows[0].length; col++) {
				s.push(table.Data.HeaderRows[0][col].text);
			}
		}
		//row and col data
		for (var row = 0; row < table.Data.Rows.length; row++) {
			for (var col = 0; col < table.Data.Rows[row].length; col++) {
				try {
					s.push(table.Data.Rows[row][col].text);
				} catch (e) {
					// text is null sometimes - not sure if it matters
					//alert(e + " " + table.Layout.Caption.label);
					s.push('MISSING');  // push empty string or the structure is lost
				};
			}
		}
		return s.join('|') + "|";
	};

	function updateArray(name, data) {
		for (var i = 0; i < objs.length; i++) {
			if (objs[i].name == name) {
				objs[i].data = data;
				break;
			}
		}
	};

	let GetObject = function (doc, objectID) {
		return new Promise(function (resolve) {
			//console.log("getting object", objectID);
			doc.GetObject(objectID, function () {
				this.DocumentMgr.RemoveFromManagers(this.ObjectMgr);
				this.DocumentMgr.RemoveFromMembers(this.ObjectMgr);
				resolve(this);
			});
		});
	};

	let GetObjectFullDataPage = async function (doc, objectID) {
		console.log("getting GetObjectFullDataPage", objectID);
		let tableObj = await GetObject(doc, objectID);
		try {
			tableObj.Data.SetPagesize(tableObj.Data.TotalSize);
		} catch (e) {
			console.log("Failed to set page size")
        }
		console.log("done SetPagesize", objectID);
		return GetObject(doc, objectID);
	};

	//let setSize = function (doc, objectID) {
	//	return new Promise(function (resolve) {
	//		doc.GetObject(objectID, function () {
	//			console.log("objectID", objectID);
	//			if (this.Type == "CH") {
	//				console.log("Set table size: " + this.Name, this.Data.TotalSize);
	//				this.Data.SetPagesize(this.Data.TotalSize);
	//			}
	//			this.DocumentMgr.RemoveFromManagers(this.ObjectMgr);
	//			this.DocumentMgr.RemoveFromMembers(this.ObjectMgr);
	//			resolve(this);
	//		});
	//	});
	//};

	function postData() {
		console.log("Posting form");
		var payload = new FormData();
		objs.forEach(o => {
			payload.append(o.fieldId, o.data)
		});
		var xhr = new XMLHttpRequest();

		xhr.onreadystatechange = function () {
			if (xhr.readyState === 4) {

				let blob = new Blob([xhr.response], { type: 'application/pdf' });
				let elem = window.document.createElement('a');
				elem.href = window.URL.createObjectURL(blob);
				elem.download = "SalesStatusOverview.pdf";
				document.body.appendChild(elem);
				elem.click();
				window.URL.revokeObjectURL(elem.href);
				document.body.removeChild(elem);
				btn.value = "Generate SSO Sheet";
				btn.disabled = false;
			}
		}

		xhr.open("POST", aspNetEndpoint, true);
		xhr.responseType = 'blob';
		xhr.send(payload);
	}

	function timeout(ms) {
		return new Promise(resolve => setTimeout(resolve, ms));
	}

	let updateData = async function (doc) {
		//debugger;
		
		//Create array of all object ids to monitor
		objs.push({ name: 'Document.Document\\CS01', data: '', fieldId: 'tblCurrentSelections' });
		objs.push({ name: 'Document.Document\\TX44', data: '', fieldId: 'txtCurrentWeek' });
		objs.push({ name: 'Document.Document\\CH111', data: '', fieldId: 'tblBuildStatus' }); // was CH63 - moved to "hidden" version as a workaround
		objs.push({ name: 'Document.Document\\CH78', data: '', fieldId: 'tblSalesStatus' });
		objs.push({ name: 'Document.Document\\CH107', data: '', fieldId: 'tblPlotValuesByPeriod' });
		objs.push({ name: 'Document.Document\\CH49', data: '', fieldId: 'tblStockStatus' });
		objs.push({ name: 'Document.Document\\CH89', data: '', fieldId: 'tblSecuredPlotCounts' });
		objs.push({ name: 'Document.Document\\CH79', data: '', fieldId: 'tblSalesStatusBudget' });

		const alldone = await objs.map(async tbl => {
			await timeout(50);
var tableObj;
if(tbl.name.substring(1, 2)==='CH') {
			tableObj = await GetObjectFullDataPage(doc, tbl.name);
} else {
	tableObj = await GetObject(doc, tbl.name);
}
			await timeout(50);
			let ts = tableToString(tableObj);
			await new Promise((resolve) => { updateArray(tbl.name, ts); resolve(this); });
			await timeout(50);
			return tableObj;
		});
		const allfinished = await Promise.all(alldone);
		console.log(allfinished);
		await new Promise((resolve) => setTimeout(function () { postData(); resolve(this) }, 100));
		//setTimeout(function () { postData() }, 100);
	}

	//let setTableSizes = async function () {
	//	const alldone = await objs.map(async tbl => {
	//		const tableObj = await setSize(doc, tbl.name);
	//		return tableObj;
	//	});
	//	const allfinished = await Promise.all(alldone);
	//	console.log("set sizes", allfinished);
	//}

	//doc = Qv.GetCurrentDocument();

	// hook submit action on form so we populate details before POST
	var btn = document.getElementById('btn');
	btn.addEventListener("click", function () {
		console.log("Button clicked");
		btn.value = "Please wait...";
		btn.disabled = true;
		updateData(Qv.GetCurrentDocument());
	});
	btn.disabled = false;
	// required otherwise data is truncated in larger datasets at 40 rows
	//await setTableSizes();


}());
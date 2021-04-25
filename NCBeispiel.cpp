#include "NCBeispiel.h"

class NCBeispiel {
private:
	int g_DBStatus = 0;
	IWebBrowser2 *g_Browser = 0;
	sqlite3 *g_db = 0;
	VARIANT vtRes;

public:
	NCBeispiel() {
		CoInitialize(0);
		VariantInit(&vtRes);

		g_DBStatus = sqlite3_open("ncdb", &g_db);

		if (g_DBStatus != SQLITE_OK)
			return;

		// Creates a table, in case it does not yet exist
		sqlite3_exec(g_db, 
			"CREATE TABLE IF NOT EXISTS NC(name TEXT, value TEXT)", 
			0, 0, 0);

		// In Windows 10 it automatically redirects to Edge, so avoid that
		DisableIERedirection();

		// Initialize an IE section
		IWB2SendRequest("about:blank", 0, 0, &g_Browser, 1000, TRUE);

		if (g_Browser)
			loadUI();
	}

	bool getFormField(const char *id, std::string& value) {
		std::ostringstream script;
		script << "document.getElementById('" << id << "').value";

		// Execute the script to retrieve the value of the form
		IWB2ExecScript(&g_Browser, script.str().c_str(), 1000, &vtRes);

		if (vtRes.vt != VT_BSTR)
			return false;

		// IWebBrowser functions return strings in unicode format,
		// so convert it to ascii for compatibility
		std::wstring_convert<std::codecvt_utf8<wchar_t>> cvt_utf8;
		const std::wstring wide_string = vtRes.bstrVal;

		value.assign(cvt_utf8.to_bytes(wide_string));
		return true;
	}

	bool isExit() {
		// Execute the script to retrieve the value of the form
		IWB2ExecScript(&g_Browser, "document.getElementById('exitSignal').value", 1000, &vtRes);
		return (vtRes.vt == VT_BSTR && vtRes.bstrVal[0] == '1') ? true : false;
	}

	bool loadUI() {
		const wchar_t *html = 
		   L"<html>"
			"<style>"
			"table,th,td{border:1px solid black;width:400px}"
			"</style>"
			"<body>"
			"<center>"
			"<img src=\"https://www.neurocheck.com/ncwp-content/uploads/neurocheck-logo-1.svg\" width=\"250\" height=\"250\">"
			"<br>"
			"<input type=\"text\" id = \"nameForm\" placeholder = \"Enter new name\">"
			"<input type=\"text\" id = \"valueForm\" placeholder = \"Enter new value\">"
			"<input type=\"hidden\" id = \"newName\">"
			"<input type=\"hidden\" id = \"newValue\">"
			"<input type=\"hidden\" id = \"exitSignal\">"
			"<br>"
			"<br>"
			"<button onclick=\"document.getElementById('newName').value=document.getElementById('nameForm').value;document.getElementById('newValue').value=document.getElementById('valueForm').value;\">"
			"Add to SQL table"
			"</button>"
			"<button onclick=\"document.getElementById('exitSignal').value=1\">"
			"Exit"
			"</button>"
			"<br>"
			"<br>"
			"<br>"
			"<br>"
			"<div id=\"db\">"
			"</div>"
			"</center>"
			"</body>"
			"</html>";

		if (!IWB2SetFileData(&g_Browser, html))
			return false;

		// List the entries in the SQL table and add them to HTML table
		sqlite3_stmt *stmt = 0;
		g_DBStatus = sqlite3_prepare_v2(g_db,
			"SELECT name, value FROM NC",
			-1, &stmt, 0);
		
		if (g_DBStatus != SQLITE_OK)
			return false;

		std::ostringstream entries;

		while ((g_DBStatus = sqlite3_step(stmt)) == SQLITE_ROW) {
			const char *name = reinterpret_cast<const char *>(sqlite3_column_text(stmt, 0));
			const char *value = reinterpret_cast<const char *>(sqlite3_column_text(stmt, 1));

			entries << "<tr><td>" << name << "</td>" << "<td>" << value << "</td></tr>";
		}

		sqlite3_finalize(stmt);
		std::ostringstream table;
		table << 
			"<table>"
			"<caption><b>*** TABLE VALUES ***</b></caption>"
			"<tr><th>Name</th><th>Value</th></tr>" 
			<< entries.str() << "</table>";

		std::ostringstream script;
		script << "document.getElementById('db').innerHTML='" << table.str() << "'";
		return IWB2ExecScript(&g_Browser, script.str().c_str(), 1000, &vtRes) ? true : false;
	}

	int NCLoop() {
		int status = 0;

		while (1) {
			// Check if the 'Exit' button was clicked
			if (isExit()) {
				break;
			}
			
			// Check if a new element should be added
			std::string newName, newValue;
			getFormField("newName", newName);
			getFormField("newValue", newValue);

			if (newName.length() && newValue.length()) {
				// Add the new entry to the SQL table
				std::ostringstream cmd;
				cmd << "INSERT INTO NC (name, value) VALUES('" <<
					newName << "', '" << newValue << "')";
				sqlite3_exec(g_db, cmd.str().c_str(), 0, 0, 0);

				// Reload the page
				loadUI();
			}

			std::this_thread::sleep_for(std::chrono::milliseconds(100));
		}
		
		return status;
	}

	~NCBeispiel() {
		// Cleanup
		VariantClear(&vtRes);
		if (g_Browser) {
			g_Browser->Quit();
			g_Browser->Release();
		}
		if (g_db) sqlite3_close(g_db);
		CoUninitialize();
	}
};

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd)
{
	NCBeispiel nc;
	return nc.NCLoop();
}

import fs from 'fs';
import * as XLSX from 'xlsx/xlsx.mjs';
import employees from './data-sources/employee.js';

class IndexSvc {

    setFsToXlsx() {
        XLSX.set_fs(fs);
    }

    async exportArrayToExcel(items, output) {
        console.log('Exporting to Excel...🚀');

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(items);
        XLSX.utils.book_append_sheet(wb, ws, 'Employees');
        XLSX.writeFile(wb, output);

        console.log('Exported to Excel successfully! ✅');
    }

    static run() {
        const items = employees;
        const OUTPUT = './output'
        const FILENAME = 'employees.xls';
        const output = `${OUTPUT}/${FILENAME}`;

        // check if output directory exists
        if (!fs.existsSync(OUTPUT)) {
            fs.mkdirSync(OUTPUT);
        }

        this.prototype.setFsToXlsx();
        this.prototype.exportArrayToExcel(items, output);
    }
}

export default IndexSvc;
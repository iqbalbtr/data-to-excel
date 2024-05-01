import { type Request, type Response, type NextFunction } from "express";
import excelJs from "exceljs";
import Excel from "../libs/excel";

const sampleData = [
    {
        "No": 1,
        "Waktu": "08:00:00",
        "Beban (kw)": 100,
        "% CAP": 80,
        "Frek (Hz)": 50,
        "PF": 0.9,
        "Suhu Air": 30,
        "Tekanan Oli (Bar)": 5,
        "RPM": 900,
        "BATT": 12,
        "KET": "OK",
        "IR (A)": 50,
        "IS (A)": 50,
        "IT (A)": 50,
        "I Avr": 50,
        "R-S": 230,
        "R-T": 230,
        "S-T": 230,
        "V Avr": 230
    },
    {
        "No": 2,
        "Waktu": "08:05:00",
        "Beban (kw)": 110,
        "% CAP": 88,
        "Frek (Hz)": 49,
        "PF": 0.85,
        "Suhu Air": 32,
        "Tekanan Oli (Bar)": 4.5,
        "RPM": 920,
        "BATT": 11,
        "KET": "LOW",
        "IR (A)": 55,
        "IS (A)": 55,
        "IT (A)": 55,
        "I Avr": 55,
        "R-S": 231,
        "R-T": 231,
        "S-T": 231,
        "V Avr": 231
    },
    {
        "No": 3,
        "Waktu": "08:10:00",
        "Beban (kw)": 105,
        "% CAP": 84,
        "Frek (Hz)": 51,
        "PF": 0.88,
        "Suhu Air": 31,
        "Tekanan Oli (Bar)": 4.8,
        "RPM": 910,
        "BATT": 10,
        "KET": "OK",
        "IR (A)": 52,
        "IS (A)": 52,
        "IT (A)": 52,
        "I Avr": 52,
        "R-S": 229,
        "R-T": 229,
        "S-T": 229,
        "V Avr": 229
    },
    {
        "No": 4,
        "Waktu": "08:15:00",
        "Beban (kw)": 95,
        "% CAP": 76,
        "Frek (Hz)": 52,
        "PF": 0.92,
        "Suhu Air": 29,
        "Tekanan Oli (Bar)": 5.2,
        "RPM": 880,
        "BATT": 12,
        "KET": "OK",
        "IR (A)": 48,
        "IS (A)": 48,
        "IT (A)": 48,
        "I Avr": 48,
        "R-S": 228,
        "R-T": 228,
        "S-T": 228,
        "V Avr": 228
    },
    {
        "No": 5,
        "Waktu": "08:20:00",
        "Beban (kw)": 120,
        "% CAP": 96,
        "Frek (Hz)": 48,
        "PF": 0.80,
        "Suhu Air": 35,
        "Tekanan Oli (Bar)": 4.0,
        "RPM": 930,
        "BATT": 11,
        "KET": "LOW",
        "IR (A)": 60,
        "IS (A)": 60,
        "IT (A)": 60,
        "I Avr": 60,
        "R-S": 232,
        "R-T": 232,
        "S-T": 232,
        "V Avr": 232
    },
    {
        "No": 6,
        "Waktu": "08:25:00",
        "Beban (kw)": 115,
        "% CAP": 92,
        "Frek (Hz)": 49,
        "PF": 0.82,
        "Suhu Air": 34,
        "Tekanan Oli (Bar)": 4.2,
        "RPM": 925,
        "BATT": 10,
        "KET": "OK",
        "IR (A)": 58,
        "IS (A)": 58,
        "IT (A)": 58,
        "I Avr": 58,
        "R-S": 230,
        "R-T": 230,
        "S-T": 230,
        "V Avr": 230
    },
    {
        "No": 7,
        "Waktu": "08:30:00",
        "Beban (kw)": 125,
        "% CAP": 100,
        "Frek (Hz)": 47,
        "PF": 0.78,
        "Suhu Air": 36,
        "Tekanan Oli (Bar)": 3.8,
        "RPM": 940,
        "BATT": 12,
        "KET": "LOW",
        "IR (A)": 62,
        "IS (A)": 62,
        "IT (A)": 62,
        "I Avr": 62,
        "R-S": 233,
        "R-T": 233,
        "S-T": 233,
        "V Avr": 233
    },
    {
        "No": 8,
        "Waktu": "08:35:00",
        "Beban (kw)": 130,
        "% CAP": 104,
        "Frek (Hz)": 46,
        "PF": 0.75,
        "Suhu Air": 38,
        "Tekanan Oli (Bar)": 3.6,
        "RPM": 950,
        "BATT": 11,
        "KET": "LOW",
        "IR (A)": 65,
        "IS (A)": 65,
        "IT (A)": 65,
        "I Avr": 65,
        "R-S": 235,
        "R-T": 235,
        "S-T": 235,
        "V Avr": 235
    },
    {
        "No": 9,
        "Waktu": "08:40:00",
        "Beban (kw)": 110,
        "% CAP": 88,
        "Frek (Hz)": 49,
        "PF": 0.85,
        "Suhu Air": 32,
        "Tekanan Oli (Bar)": 4.5,
        "RPM": 920,
        "BATT": 10,
    }
]

interface Data {
    No: number;
    Waktu: string;
    'Beban (kw)': number;
    '% CAP': number;
    'Frek (Hz)': number;
    PF: number;
    'Suhu Air': number;
    'Tekanan Oli (Bar)': number;
    RPM: number;
    BATT: number;
    KET: string;
    'IR (A)': number;
    'IS (A)': number;
    'IT (A)': number;
    'I Avr': number;
    'R-S': number;
    'R-T': number;
    'S-T': number;
    'V Avr': number;
}


async function excel(req: Request, res: Response, next: NextFunction) {
    try {

        const excel = new excelJs.Workbook();

        const excel_test = new Excel.Workbook("Data");

        excel_test.titleRow({
            value: "Logo",
            position: "A1:B4"
        });
        excel_test.titleRow({
            value: "Logsheet",
            position: "C1:S1"
        });


        excel_test.headRow({
            value: "Hari tangal",
            position: "C2:H2"
        });
        excel_test.headRow({
            value: "Hari",
            position: "I2:N2"
        });
        excel_test.headRow({
            value: "Pemilik",
            position: "O2:S2"
        });


        excel_test.headRow({
            value: "Mesin",
            position: "C3:H3"
        });
        excel_test.headRow({
            value: "Model",
            position: "I3:N3"
        });
        excel_test.headRow({
            value: "Pemilik",
            position: "O3:S3"
        });

        excel_test.headRow({
            value: "Mesin",
            position: "C4:H4"
        });
        excel_test.headRow({
            value: "Model",
            position: "I4:N4"
        });
        excel_test.headRow({
            value: "Pemilik",
            position: "O4:S4"
        });

        excel_test.headRow({
            value: "Arus",
            position: "E5:H5"
        });
        excel_test.headRow({
            value: "Tegangan",
            position: "I5:L5"
        });


        excel_test.headRow({
            value: "No",
            center: true,
            position: "A5:A6"
        })

        excel_test.headRow({
            value: "Waktu",
            center: true,
            position: "B5:B6"
        })
        excel_test.headRow({
            value: "Beban (kw)",
            center: true,
            position: "C5:C6"
        })
        excel_test.headRow({
            value: "% CAP",
            center: true,
            position: "D5:D6"
        })
        excel_test.headRow({
            value: "IR (A)",
            center: true,
            position: "E6"
        })
        excel_test.headRow({
            value: "IS (A)",
            center: true,
            position: "F6"
        })
        excel_test.headRow({
            value: "IT (A)",
            center: true,
            position: "G6"
        })
        excel_test.headRow({
            value: "I Avr",
            center: true,
            position: "H6"
        })
        excel_test.headRow({
            value: "R-S",
            center: true,
            position: "I6"
        })
        excel_test.headRow({
            value: "R-T",
            center: true,
            position: "J6"
        })
        excel_test.headRow({
            value: "S-T",
            center: true,
            position: "K6"
        })
        excel_test.headRow({
            value: "V Avr",
            center: true,
            position: "L6"
        })
        excel_test.headRow({
            value: "Frek (Hz)",
            center: true,
            position: "M5:M6"
        })
        excel_test.headRow({
            value: "PF",
            center: true,
            position: "N5:N6"
        })
        excel_test.headRow({
            value: "Suhu Air",
            center: true,
            position: "O5:O6"
        })
        excel_test.headRow({
            value: "Tekanan Oli (Bar)",
            center: true,
            position: "P5:P6"
        })
        excel_test.headRow({
            value: "RPM",
            center: true,
            position: "Q5:Q6"
        })
        excel_test.headRow({
            value: "BATT",
            center: true,
            position: "R5:R6"
        })
        excel_test.headRow({
            value: "KET",
            center: true,
            position: "S5:S6"
        })

        const payload = sampleData.map(item => {
            const payloadArray: any[] = [];
            Object.values(item).forEach(itemData => {
                payloadArray.push(itemData);
            })
            return payloadArray;
        })

        excel_test.dataRow<any>(payload);

        const buffer = await excel_test.write();

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadworksheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=data.xlsx');
        res.send(buffer);
    } catch (error) {

    }
}

async function get(req: Request, res: Response, next: NextFunction) {
    try {
        res.sendFile("index.html")
    } catch (error) {

    }
}

export default {
    get,
    excel
}

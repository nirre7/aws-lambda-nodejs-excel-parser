import {GetObjectCommand, S3Client} from '@aws-sdk/client-s3'
import * as XLSX from 'xlsx'
import {Readable} from 'stream'

const s3 = new S3Client({region: 'eu-north-1'})

export const handler = async (event) => {

    const s3Response = await s3.send(new GetObjectCommand({
        Bucket: "nisoexcelbucket",
        Key: "test.xlsx"
    }))

    const file = await s3Response.Body as Readable
    const buffer = await readFile(file)
    const sheets = await parseWorkBook(buffer)

    const response = {
        statusCode: 200,
        body: JSON.stringify('Done parsing excel'),
    };
    return response;
};

function readFile(file: (Readable)): Promise<Buffer> {
    return new Promise(resolve => {
        const bufs = [];
        (file as Readable).on('data', function (data) {
            bufs.push(data);
        });
        (file as Readable).on('end', function () {
            resolve(Buffer.concat(bufs))
        });
    })
}

function parseWorkBook(buffer: Buffer) {
    const workBook = XLSX.read(buffer)

    let sheet = workBook.Sheets[workBook.SheetNames[0]];

    // const data: any[] = XLSX.utils.sheet_to_json(sheet, {});

    var range = XLSX.utils.decode_range(sheet['!ref']);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        // Example: Get second cell in each row, i.e. Column "B"
        const firstCell = sheet[XLSX.utils.encode_cell({r: rowNum, c: 0})];
        // NOTE: firstCell is undefined if it does not exist (i.e. if its empty)
        console.log(firstCell); // firstCell.v contains the value, i.e. string or number
    }

    // console.log(data)

    return new Promise(resolve => {
        resolve(workBook.SheetNames.join(', '))
    })
}


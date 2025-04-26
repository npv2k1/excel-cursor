# Excel Cursor

Thư viện TypeScript giúp thao tác với file Excel một cách dễ dàng thông qua API dạng con trỏ (cursor).

## Cài đặt

```bash
npm install excel-cursor
# hoặc
yarn add excel-cursor
```

## Tính năng

- Di chuyển con trỏ linh hoạt trong file Excel
- Đọc/ghi dữ liệu vào ô
- Định dạng ô (font, màu sắc, căn lề)
- Merge ô
- Thêm/xóa hàng
- Chuyển đổi giữa các worksheet
- Hỗ trợ công thức Excel
- Định dạng có điều kiện

## Sử dụng

```typescript
import { Workbook } from 'exceljs';
import { ExcelCursor } from 'excel-cursor';

// Khởi tạo workbook và cursor
const workbook = new Workbook();
const cursor = new ExcelCursor(workbook);

// Di chuyển và nhập dữ liệu
cursor
  .move('A1')
  .setData('Hello')
  .nextRow()
  .setData('World')
  .formatCell({
    font: { bold: true },
    alignment: { vertical: 'middle', horizontal: 'center' }
  });

// Lưu file
await workbook.xlsx.writeFile('output.xlsx');
```

## API

### Di chuyển con trỏ

- `move(address: string)`: Di chuyển đến địa chỉ ô cụ thể (vd: 'A1')
- `moveTo(row: number, col: number)`: Di chuyển đến vị trí hàng và cột
- `nextRow(n = 1)`: Di chuyển xuống n hàng
- `prevRow(n = 1)`: Di chuyển lên n hàng
- `nextCol(n = 1)`: Di chuyển sang phải n cột
- `prevCol(n = 1)`: Di chuyển sang trái n cột

### Thao tác dữ liệu

- `setData(data: any, address?: string)`: Ghi dữ liệu vào ô hiện tại hoặc ô chỉ định
- `setFormula(formula: string, address?: string)`: Đặt công thức cho ô
- `formatCell(format: CellFormat, address?: string)`: Định dạng ô

### Định dạng

- `colSpan(n: number, address?: string)`: Merge n cột từ vị trí hiện tại
- `rowSpan(n: number, address?: string)`: Merge n hàng từ vị trí hiện tại
- `setColWidth(width: number, colOrAddress?: number | string)`: Đặt độ rộng cột
- `setRowHeight(height: number, rowOrAddress?: number | string)`: Đặt chiều cao hàng

### Worksheet

- `switchSheet(sheetName: string)`: Chuyển sang worksheet khác
- `createSheet(sheetName: string)`: Tạo worksheet mới

## Đóng góp

Mọi đóng góp đều được hoan nghênh! Vui lòng:

1. Fork dự án
2. Tạo branch cho tính năng (`git checkout -b feature/amazing-feature`)
3. Commit thay đổi (`git commit -m 'Add some amazing feature'`)
4. Push lên branch (`git push origin feature/amazing-feature`)
5. Tạo Pull Request

## Giấy phép

Dự án này được phân phối dưới giấy phép MIT. Xem file `LICENSE` để biết thêm chi tiết.

## Tác giả

Nguyễn Văn A - [@nguyenvana](https://github.com/nguyenvana)

## Hỗ trợ

Nếu bạn gặp vấn đề hoặc có câu hỏi, vui lòng tạo issue trong repository.
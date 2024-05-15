import { connect, query, close } from 'mssql';

// Cấu hình kết nối
const config = {
    user: 'admin',      // Tên người dùng SQL Server
    password: '123456',  // Mật khẩu SQL Server
    server: '10.0.0.18',   // Địa chỉ server, có thể là 'localhost' hoặc địa chỉ IP
    database: 'hrdb',  // Tên database bạn muốn kết nối
//                options: {
//                  encrypt: true,     // Dành cho Azure
//                enableArithAbort: true
//          }
};

// Kết nối đến SQL Server và thực hiện query
async function connectAndQuery() {
    try {
        // Kết nối đến SQL Server
        await connect(config);

        // Thực hiện query
        const result = await query`SELECT * FROM hrdb.dbo.congnhanvien`;

        console.log(result);

    } catch (err) {
        console.error('Error:', err);
    } finally {
        // Đóng kết nối
        close();
    }
}

// Gọi hàm connectAndQuery để thực hiện
connectAndQuery();
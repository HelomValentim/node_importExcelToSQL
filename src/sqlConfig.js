const config = {
    user: 'my_user',
    password: 'my_password',
    server: 'my_server_addres',
    database: 'my_database_name',
    port: 1234,
    options: {
        trustServerCertificate: true,
        connectionTimeout: 60000,
    }
};

module.exports = config;
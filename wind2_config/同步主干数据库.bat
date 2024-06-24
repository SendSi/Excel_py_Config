del mainDBData.sql

mysqldump -h192.168.0.239 -P3306 -uplan -p123456 wind_base_config>mainDBData.sql

mysql -h192.168.0.239 -P3306 -uplan -p123456 wind_cehua_lujun_config<mainDBData.sql

del mainDBData.sql
pause
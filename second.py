import configparser
config = configparser.ConfigParser()
print(config.read('settings.ini'))
print("Sections : ", config.sections())
print("Installation Library : ", config.get('final_name_file', 'final_file'))
print("Log Errors debugged ? : ", config.getboolean('debug', 'log_errors'))
print("Port Server : ", config.getint('server', 'port'))
for value in config['files_sheet'].values():
    print(value)




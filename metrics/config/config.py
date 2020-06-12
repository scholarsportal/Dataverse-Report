from configparser import ConfigParser


def read(section, filename='metrics/config/metrics.cfg'):
    parser = ConfigParser()
    parser.read(filename)
    cfg = {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            cfg[param[0]] = param[1]
    else:
        raise Exception('Section {0} not found in the {1} file'.format(section, filename))
    return cfg

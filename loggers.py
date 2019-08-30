import logging

class Loggers():
    def logging(self, namelogger, level):
        logger = logging.getLogger(namelogger)
        logger.setLevel(level)

        # create the logging file handler
        fh = logging.FileHandler("ActUnFixBarCode.log")
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)

        # add handler to logger object
        logger.addHandler(fh)
        return logger
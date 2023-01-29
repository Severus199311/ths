import logging
from logging import handlers
import os
import sys

class logger():
	log_path = 'D:/code/logs'
	log_size = 1024*1024*5

	logname = os.path.join(log_path, sys.argv[0].split('/')[-1].split('.')[0])
	log = logging.getLogger('')
	fmt = logging.Formatter('[%(asctime)s][%(levelname)s] %(message)s', '%Y-%m-%d %H:%M:%S')
	handle_a = logging.handlers.RotatingFileHandler(logname, maxBytes=log_size, backupCount=2, encoding="utf-8")
	handle_a.setFormatter(fmt)
	log.addHandler(handle_a)
	handle_b = logging.StreamHandler(stream=sys.stdout)
	handle_b.setFormatter(fmt)
	log.addHandler(handle_b)

	log.setLevel(logging.INFO)

	@classmethod
	def info(cls, msg):
		cls.log.info(msg)
		return

	@classmethod
	def error(cls, msg):
		cls.log.error(msg)
		return

	@classmethod
	def critical(cls, msg):
		cls.log.critical(msg)
		return
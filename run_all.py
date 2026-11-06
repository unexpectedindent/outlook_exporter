"""Execute both exports and rebuild derived thread text."""
import argparse
import logging
from common import load_config
from export_mail import run as mail
from export_calendar import run as calendar
from thread_text import process

if __name__ == '__main__':
    p=argparse.ArgumentParser()
    p.add_argument('--config', default='config.json')
    args=p.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    cfg=load_config(args.config)
    mail(cfg)
    calendar(cfg)
    process(cfg)

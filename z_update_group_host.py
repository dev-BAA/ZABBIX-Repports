# -*- coding: utf-8 -*-
from datetime import datetime
import os, sys
import time
import calendar
import datetime
import openpyxl
import logging
sys.path.append("/root")
from pyzabbix import ZabbixAPI
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook
from settings import *

logging.basicConfig(level=logging.ERROR,
                    format='%(asctime)s %(name)s %(levelname)s %(message)s',
                    datefmt='%d-%m %H:%M:%S',
                    filename='./logs/log_update_group_host')
logger = logging.getLogger('-')
logger.setLevel(logging.INFO)

server = 'http://127.0.0.1/zabbix'
zapi = ZabbixAPI(server)
zapi.login(login, pswd)

logger.info("=======================================" + time.ctime() + "=======================================")
o = 0
hosts = zapi.host.get(groupids=[11,12,14,19],
                      #output=["hostid", "name"]
                      output='extend'
                      )

for point in hosts:
 logger.info("-------------------------------------")
 hostname = point['name']
 hid = point['hostid']
 o+=1
 print(o)
 logger.info(o)
 logger.info(hostname)
 logger.info(hid)
 if hostname[0:2] == "U-":
   iten = zapi.item.get(hostids=hid,
                        output=["lastvalue", "name"],
                        filter={'name':'Модель'}
                        )
   mdel = iten[0]['lastvalue']
   logger.info(mdel)
   grups = zapi.hostgroup.get(hostids=hid,
                      output='extend'
                      )
   logger.info("**************************")
   logger.info(len(grups))
   if len(grups) > 2:
     logger.info("-----")
     logger.info(point)
     logger.info("-----")
     logger.info(iten)
     logger.info("-----")
     logger.info(mdel[0:2])
     if mdel[0:2] == "FS" or mdel[0:2] == "EC":
       zapi.host.update(hostid=hid,groups=['5','11'],templates=['10379'])
     elif mdel[0:2] == "HP":
       zapi.host.update(hostid=hid,groups=['5','12'],templates=['10528'])
     elif mdel[0:2] == "Ca":
       zapi.host.update(hostid=hid,groups=['5','14'],templates=['10576'])
     elif mdel[0:2] == "Xe":
       zapi.host.update(hostid=hid,groups=['5','19'],templates=['10788'])
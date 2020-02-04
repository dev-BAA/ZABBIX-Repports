# -*- coding: utf-8 -*-
import os, sys
import time
import calendar
import datetime
import openpyxl
import logging
sys.path.append("/root")
from pyzabbix import ZabbixAPI
from datetime import datetime
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook
from settings import *

logging.basicConfig(level=logging.ERROR,
                    format='%(asctime)s %(name)s %(levelname)s %(message)s',
                    datefmt='%d-%m %H:%M:%S',
                    filename='./logs/log_new_host')
logger = logging.getLogger('-')
logger.setLevel(logging.INFO)

server = 'http://127.0.0.1/zabbix'
zapi = ZabbixAPI(server)
zapi.login(login, pswd)

logger.info("=======================================" + time.ctime() + "=======================================")
o = 0
hosts = zapi.host.get(groupids=[11,12,14,19],
                      output='extend'
                      )
for point in hosts:
  point_name = point['name']
  point_hostid = point['hostid']
  o+=1
  print(o)
  logger.info("------------------------------------------------------------------------------------------")
  logger.info(o)
  logger.info(time.ctime())
  logger.info(point_name)
  logger.info(point_hostid)
  if "ipoe-users" in point_name:
    logger.info("------------------------")
    logger.info(point)
    point_interface = zapi.hostinterface.get(hostids=point_hostid,
                                 output='extend')
    logger.info(point_interface)
    logger.info("IP адрес обнаруженного хоста")
    point_ip = point_interface[0]['ip']
    logger.info(point_ip)
    logger.info("------------------------")
    ### Получение элементов данных обнаруженного хоста
    point_items = zapi.item.get(hostids=point_hostid,
                         output=["lastvalue", "name"],
                         filter={'name':'Имя'}
                         )
    point_nameU = point_items[0]['lastvalue']
    logger.info("Имя обнаруженного хоста")
    logger.info(point_nameU)
    logger.info("------------------------")
    ### Проверка указанного имени хоста
    if point_nameU != 0 and point_nameU[0:2] == "U-":
      ### Получение хостов с именем обнаруженного хоста (если такие есть)
      hostsU = zapi.host.get(groupids=[11,12,14,19],
                             output=["hostid", "name"],
                             filter={'name':point_nameU}
                             )
      logger.info("------------------------")
      logger.info(len(hostsU))
      ## НЕТ ХОСТОВ С ТАКИМ ИМЕНЕМ
      if len(hostsU) == 0:
        logger.info("hostsU == 0, нет хостов с таким именем")
        zapi.host.update(hostid=point_hostid, host=point_nameU, name=point_nameU)
      ## ЕСТЬ ХОСТ С ТАКИМ ИМЕНЕМ
      elif len(hostsU) != 0:
        logger.info("hostsU != 0, хосты с таким именем есть")
        hostsU_hostid = hostsU[0]['hostid']
        hostsU_name = hostsU[0]['name']
        logger.info("Получение name существующего хоста")
        logger.info(hostsU_name)
        logger.info("Получение hostid существующего хоста")
        logger.info(hostsU_hostid)
        ### Получение hostinterface существующего хоста
        hostsU_hostiface = zapi.hostinterface.get(hostids=[hostsU_hostid],
                                                  output='extend'
                                                  )
        logger.info("Сетевой интерфейс существующего хоста")
        logger.info(hostsU_hostiface)
        ### ВАРИАНТ КОГДА У УСТРОЙСТВА ПОМЕНЯЛСЯ IP АДРЕСС
        if hostsU_name == point_nameU:
          logger.info("Новый ip адрес у существующего хоста, имеющего ip")
          logger.info(point_ip)
          hif = hostsU_hostiface[0]['interfaceid']
          logger.info("id сетевого интерфейса существующего хоста")
          logger.info(hif)
          zapi.hostinterface.update(interfaceid=hif, ip=point_ip)
          zapi.host.delete(point_hostid)
    logger.info("------------------------")
  ## ЕСЛИ ИМЯ ХОСТА СООТВЕТСТВУЮЩЕГО ФОРМАТА ПРОВЕРИМ СООТВЕТСТВУЕТ ЛИ ОНО ПАРАМЕТРУ sysName, ЕСЛИ НЕ СООТВЕТСТВУЕТ - ПЕРЕИМЕНОВЫВАЕМ ХОСТ
  else:
    logger.info("Хост уже есть")
    item_name = zapi.item.get(hostids=[point_hostid],
                         output=["lastvalue"],
                         filter={'name':'Имя'}
                         )
    item_name_new = item_name[0]['lastvalue']
    logger.info(point_name)
    logger.info(item_name_new)
    if (item_name_new.find("U-") != -1) and (item_name_new != point_name):
      logger.info("point_name не равно item_name_new")
      try:
        zapi.host.update(hostid=point_hostid, host=item_name[0]['lastvalue'], name=item_name[0]['lastvalue'])
        logger.info("Имя хоста обновлено через его sysName")
      except BaseException:
        logger.info("Не получилось обновить имя хоста через его sysName")
    if ((item_name_new == "0") or (item_name_new == "")):
      logger.info("############################################################### - sysName нужно проверить")
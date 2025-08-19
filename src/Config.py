from collections import namedtuple
import yaml as yl
import os

class Settings:
    def Ler():
        settings={}
        for a in os.listdir('settings'):
            with open(f'settings/{a}',encoding='utf8') as file:
                settings.update({a.split('.')[0]: yl.safe_load(file)})
        
        for c in settings:
            settings[c] = __class__.CreateTuples('config',settings[c])

        settings = __class__.CreateTuples('setup',settings)
        return settings

    def CreateTuples(name,args):
        obj = namedtuple(name,args)
        return obj(**args)

# Veeam-Job-Stat
Модуль **[Veeam-Job-Stat.psm1](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Veeam-Job-Stat/Veeam-Job-Stat.psm1)** для сбора и вывода статистики всех заданий Backup в **CustomObject**.

**Property:** \
**EnabledJob** - выключено/выключено задание \
**JobName** - Имя задания \
**VmCount** - Кол-во виртуальных машин в задачии \
**VmName** - Имена виртуальных машин в задании \
**JobType** - Тип Backup (**Backup/EpAgentBackup**) \
**LatestRunLocal** - Время последней **попытки запуска** задания \
**TimeLastCreation** - Время последнего начала выполнения задания \
**TimeLastCompletion** - Время последнего **успешного завершения** задания \
**RunTime** - Время выполнения задания (разница между TimeLastCreation и TimeLastCompletion) \
**RepositoryType** - Тип репозитория хранения (например, Windows) \
**Repository** - Имя виртуальной машины с репозиторием \
**DirPath** - Локальный путь директории на репозитории с хранящимися файлами Backup машины \
**VmSize** - Исходный размер виртуальной машины \
**BackupName** - Имя и расширение файла последней резервной копии \
**BackupType** - **Тип последней резервной копии (Full/Increment)** \
**BackupSize** - **Размер последней резервной копии**

Основывался на **[BR-Check-SLA](https://github.com/VeeamHub/powershell/tree/master/BR-Check-SLA)** из офиуиального репозитория **VeeamHub**, с целью миниммизировать код и переработать вывод.

![Image alt](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Screen/Module.jpg)
![Image alt](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Screen/Report-Script.jpg)
![Image alt](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Screen/Report-Message.jpg)

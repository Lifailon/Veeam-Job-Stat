# Veeam-Job-Stat

![Image alt](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Screen/Logo.jpg)

Модуль **[Veeam-Job-Stat.psm1](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Veeam-Job-Stat/Veeam-Job-Stat.psm1)** для сбора и вывода статистики всех заданий резервной копии в **CustomObject**.

![Image alt](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Screen/Module.jpg)

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

**Зависимости:**
* **Модуль Veeam.Backup.PowerShell**, который идет в составе с дистрибутивом **Veeam Backup & Replication**
* Для локального запуска модуля PowerShell на машине, требуется УЗ с правами доступа к консоли Veeam
* Для удаленного запуска необходимо добавить аудентификацию через **Connect-VBRServer** (в модуле не используется). Сам модуль можно установить из репозитория **Chocolatey** 

> Основывался на **[BR-Check-SLA](https://github.com/VeeamHub/powershell/tree/master/BR-Check-SLA)** из официального репозитория **[VeeamHub](https://github.com/VeeamHub)**, с целью миниммизировать код и переработать вывод.

## Скрипт [Veeam-Job-Stat-Report](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Veeam-Job-Stat-Report/Veeam-Job-Stat-Report.ps1) для отправки ежедневного отчета на почту.

* Встроен модуль [Export-Excel](https://github.com/Lifailon/RSA-Modules#export-excel) для создания Excel-таблицы.
* Создается директория в корне диска `С:\Veeam-Job-Stat-Log` (в случае ее отсутствия) для хранения ежедневных Excel-отчетов по дате.
* При первом запуске скрипта необходимо заполнить **Credential** для авторизации пользователем, из под которого будет происходить отправка почты. Файл с кредами будет сохранен в **Cred-Email.xml**, который располагается рядом с скриптом.

![Image alt](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Screen/Report-Script.jpg)

![Image alt](https://github.com/Lifailon/Veeam-Job-Stat/blob/rsa/Screen/Report-Message.jpg)

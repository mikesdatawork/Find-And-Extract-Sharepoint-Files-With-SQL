![MIKES DATA WORK GIT REPO](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_01.png "Mikes Data Work")        

# Find And Extract Sharepoint Files With SQL
**Post Date: May 10, 2018**        


## Contents    
- [About Process](##About-Process)  
- [SQL Logic](#SQL-Logic)  
- [Build Info](#Build-Info)  
- [Author](#Author)  
- [License](#License)       

## About-Process


![Extract Sharepoint Files With SQL]( https://mikesdatawork.files.wordpress.com/2018/05/image0012.png "Extract Sharepoint File With SQL")
 
<p>Here is some quick SQL logic to help you find extract, export, or otherwise download Sharepoint documents with SQL. The sp_oamethod has been around for quite some time, and delighted it's still working. You never know when MS will pull these procedures.</p>     


## SQL-Logic
```SQL
use [WSS_Content];
set nocount on
go
-- look for specific document files in sharepoint
select
    'database'  = db_name()
,   'time_created'  = left(alldocs.timecreated, 19)
,   'kb'        = (convert(bigint,alldocstreams.size))/1024
,   'mb'        = (convert(bigint,alldocstreams.size))/1024/1024
,   'list_name' = alllists.tp_title
,   'file_name' = alldocs.leafname
,   'url'       = alldocs.dirname
 
from
    alldocs join alldocstreams  on alldocs.id=alldocstreams.id 
    join alllists           on alllists.tp_id = alldocs.listid
where
    alldocstreams.[size] > 2048 -- looking for files under 2 mb
    and [alllists].[tp_title]   like '%training%'
    and [alldocs].[leafname]    like '%tutorial%'
    and right([alldocs].[leafname], 2) in ('oc', 'cx', 'df', 'sg', 'xt')
order by
    alldocs.timecreated desc
 
 
-- extract files using the leafname and dirname as identifiers
declare @object_token       int
declare @destination_path   varchar(255)
declare @content_binary     varbinary(max)
set @destination_path   = '\\MyServerName\W$\Sharepoint_Extraction\my_tutorial.pdf'
 
select  @content_binary     = alldocstreams.content from alldocs join alldocstreams on alldocs.id = alldocstreams.id join alllists on alllists.tp_id = alldocs.listid
where  
    alldocs.leafname    = 'my_tutorial.pdf'
    and alldocs.dirname = 'Site/Process/Tutorials/TechnicalDocuments/How-To/Guides'
 
exec sp_oacreate 'adodb.stream', @object_token output
exec sp_oasetproperty   @object_token, 'type', 1
exec sp_oamethod    @object_token, 'open'
exec sp_oamethod    @object_token, 'write',     null, @content_binary
exec sp_oamethod    @object_token, 'savetofile',    null, @destination_path, 2
exec sp_oamethod    @object_token, 'close'
exec sp_oadestroy   @object_token
 
-- see files at the destination
declare @check_files    varchar(255)
set     @check_files    = '\\MyServerName\W$\Sharepoint_Extraction\'
exec    master..xp_dirtree @check_files, 1,1
go

```

[![WorksEveryTime](https://forthebadge.com/images/badges/60-percent-of-the-time-works-every-time.svg)](https://shitday.de/)

## Build-Info

| Build Quality | Build History |
|--|--|
|<table><tr><td>[![Build-Status](https://ci.appveyor.com/api/projects/status/pjxh5g91jpbh7t84?svg?style=flat-square)](#)</td></tr><tr><td>[![Coverage](https://coveralls.io/repos/github/tygerbytes/ResourceFitness/badge.svg?style=flat-square)](#)</td></tr><tr><td>[![Nuget](https://img.shields.io/nuget/v/TW.Resfit.Core.svg?style=flat-square)](#)</td></tr></table>|<table><tr><td>[![Build history](https://buildstats.info/appveyor/chart/tygerbytes/resourcefitness)](#)</td></tr></table>|

## Author

[![Gist](https://img.shields.io/badge/Gist-MikesDataWork-<COLOR>.svg)](https://gist.github.com/mikesdatawork)
[![Twitter](https://img.shields.io/badge/Twitter-MikesDataWork-<COLOR>.svg)](https://twitter.com/mikesdatawork)
[![Wordpress](https://img.shields.io/badge/Wordpress-MikesDataWork-<COLOR>.svg)](https://mikesdatawork.wordpress.com/)

   
## License
[![LicenseCCSA](https://img.shields.io/badge/License-CreativeCommonsSA-<COLOR>.svg)](https://creativecommons.org/share-your-work/licensing-types-examples/)

![Mikes Data Work](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_02.png "Mikes Data Work")


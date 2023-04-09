echo create an sql script which can be run in one go.
MD All 
type header.sql > All/all.sql
type exist_procs.sql >> All/all.sql
type utilities.sql >> All/all.sql
type functions.sql >> All/all.sql
type appParameters.sql >> All/all.sql
type user_log.sql >> All/all.sql
type semaphore.sql >> All/all.sql
type app_log.sql >> All/all.sql
type data.sql >> All/all.sql
type meta_views.sql >> All/all.sql
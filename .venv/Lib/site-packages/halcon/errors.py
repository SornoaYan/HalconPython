# Note: According to the PEP8, constants are usually defined
# on a module level and written in all capital letters with
# underscores separating words.

# Above is a convention, it doesn't actually prevent reassignment.


#: Normal return value 
H_MSG_OK = 2
#: true 
H_MSG_TRUE = H_MSG_OK
#: false 
H_MSG_FALSE = 3
#: Stop processing 
H_MSG_VOID = 4
#: Call failed 
H_MSG_FAIL = 5
#: for internal use 
H_ERR_BREAK = 20
#: operator was canceled for hdevengine 
H_ERR_HEN_CANCEL = 21
#: operator was generally cancelled 
H_ERR_CANCEL = 22
#: for internal use 
H_ERR_TIMEOUT_BREAK = 23
#: Wrong type of control parameter: 1 
H_ERR_WIPT1 = 1201
#: Wrong type of control parameter: 2 
H_ERR_WIPT2 = 1202
#: Wrong type of control parameter: 3 
H_ERR_WIPT3 = 1203
#: Wrong type of control parameter: 4 
H_ERR_WIPT4 = 1204
#: Wrong type of control parameter: 5 
H_ERR_WIPT5 = 1205
#: Wrong type of control parameter: 6 
H_ERR_WIPT6 = 1206
#: Wrong type of control parameter: 7 
H_ERR_WIPT7 = 1207
#: Wrong type of control parameter: 8 
H_ERR_WIPT8 = 1208
#: Wrong type of control parameter: 9 
H_ERR_WIPT9 = 1209
#: Wrong type of control parameter: 10 
H_ERR_WIPT10 = 1210
#: Wrong type of control parameter: 11 
H_ERR_WIPT11 = 1211
#: Wrong type of control parameter: 12 
H_ERR_WIPT12 = 1212
#: Wrong type of control parameter: 13 
H_ERR_WIPT13 = 1213
#: Wrong type of control parameter: 14 
H_ERR_WIPT14 = 1214
#: Wrong type of control parameter: 15 
H_ERR_WIPT15 = 1215
#: Wrong type of control parameter: 16 
H_ERR_WIPT16 = 1216
#: Wrong type of control parameter: 17 
H_ERR_WIPT17 = 1217
#: Wrong type of control parameter: 18 
H_ERR_WIPT18 = 1218
#: Wrong type of control parameter: 19 
H_ERR_WIPT19 = 1219
#: Wrong type of control parameter: 20 
H_ERR_WIPT20 = 1220
#: Wrong value of control parameter: 1 
H_ERR_WIPV1 = 1301
#: Wrong value of control parameter: 2 
H_ERR_WIPV2 = 1302
#: Wrong value of control parameter: 3 
H_ERR_WIPV3 = 1303
#: Wrong value of control parameter: 4 
H_ERR_WIPV4 = 1304
#: Wrong value of control parameter: 5 
H_ERR_WIPV5 = 1305
#: Wrong value of control parameter: 6 
H_ERR_WIPV6 = 1306
#: Wrong value of control parameter: 7 
H_ERR_WIPV7 = 1307
#: Wrong value of control parameter: 8 
H_ERR_WIPV8 = 1308
#: Wrong value of control parameter: 9 
H_ERR_WIPV9 = 1309
#: Wrong value of control parameter: 10 
H_ERR_WIPV10 = 1310
#: Wrong value of control parameter: 11 
H_ERR_WIPV11 = 1311
#: Wrong value of control parameter: 12 
H_ERR_WIPV12 = 1312
#: Wrong value of control parameter: 13 
H_ERR_WIPV13 = 1313
#: Wrong value of control parameter: 14 
H_ERR_WIPV14 = 1314
#: Wrong value of control parameter: 15 
H_ERR_WIPV15 = 1315
#: Wrong value of control parameter: 16 
H_ERR_WIPV16 = 1316
#: Wrong value of control parameter: 17 
H_ERR_WIPV17 = 1317
#: Wrong value of control parameter: 18 
H_ERR_WIPV18 = 1318
#: Wrong value of control parameter: 19 
H_ERR_WIPV19 = 1319
#: Wrong value of control parameter: 20 
H_ERR_WIPV20 = 1320
#: Wrong value of component 
H_ERR_WCOMP = 1350
#: Wrong value of gray value component 
H_ERR_WGCOMP = 1351
#: Wrong number of values of ctrl.par.: 1 
H_ERR_WIPN1 = 1401
#: Wrong number of values of ctrl.par.: 2 
H_ERR_WIPN2 = 1402
#: Wrong number of values of ctrl.par.: 3 
H_ERR_WIPN3 = 1403
#: Wrong number of values of ctrl.par.: 4 
H_ERR_WIPN4 = 1404
#: Wrong number of values of ctrl.par.: 5 
H_ERR_WIPN5 = 1405
#: Wrong number of values of ctrl.par.: 6 
H_ERR_WIPN6 = 1406
#: Wrong number of values of ctrl.par.: 7 
H_ERR_WIPN7 = 1407
#: Wrong number of values of ctrl.par.: 8 
H_ERR_WIPN8 = 1408
#: Wrong number of values of ctrl.par.: 9 
H_ERR_WIPN9 = 1409
#: Wrong number of values of ctrl.par.: 10 
H_ERR_WIPN10 = 1410
#: Wrong number of values of ctrl.par.: 11 
H_ERR_WIPN11 = 1411
#: Wrong number of values of ctrl.par.: 12 
H_ERR_WIPN12 = 1412
#: Wrong number of values of ctrl.par.: 13 
H_ERR_WIPN13 = 1413
#: Wrong number of values of ctrl.par.: 14 
H_ERR_WIPN14 = 1414
#: Wrong number of values of ctrl.par.: 15 
H_ERR_WIPN15 = 1415
#: Wrong number of values of ctrl.par.: 16 
H_ERR_WIPN16 = 1416
#: Wrong number of values of ctrl.par.: 17 
H_ERR_WIPN17 = 1417
#: Wrong number of values of ctrl.par.: 18 
H_ERR_WIPN18 = 1418
#: Wrong number of values of ctrl.par.: 19 
H_ERR_WIPN19 = 1419
#: Wrong number of values of ctrl.par.: 20 
H_ERR_WIPN20 = 1420
#: Number of input objects too big 
H_ERR_IONTB = 1500
#: Wrong number of values of object par.: 1 
H_ERR_WION1 = 1501
#: Wrong number of values of object par.: 2 
H_ERR_WION2 = 1502
#: Wrong number of values of object par.: 3 
H_ERR_WION3 = 1503
#: Wrong number of values of object par.: 4 
H_ERR_WION4 = 1504
#: Wrong number of values of object par.: 5 
H_ERR_WION5 = 1505
#: Wrong number of values of object par.: 6 
H_ERR_WION6 = 1506
#: Wrong number of values of object par.: 7 
H_ERR_WION7 = 1507
#: Wrong number of values of object par.: 8 
H_ERR_WION8 = 1508
#: Wrong number of values of object par.: 9 
H_ERR_WION9 = 1509
#: Number of output objects too big 
H_ERR_OONTB = 1510
#: Wrong specification of parameter (error in file: xxx.def) 
H_ERR_WNP = 2000
#: Initialize Halcon: reset_obj_db(Width,Height,Components) 
H_ERR_HONI = 2001
#: Used number of symbolic object names too big 
H_ERR_WRKNN = 2002
#: No license found 
H_ERR_LIC_NO_LICENSE = 2003
#: No modules in license (no VENDOR_STRING) 
H_ERR_LIC_NO_MODULES = 2005
#: No license for this operator 
H_ERR_LIC_NO_LIC_OPER = 2006
#: Vendor keys do not support this platform 
H_ERR_LIC_BADPLATFORM = 2008
#: Bad vendor keys 
H_ERR_LIC_BADVENDORKEY = 2009
#: System clock has been set back 
H_ERR_LIC_BADSYSDATE = 2021
#: Version argument is invalid floating point format 
H_ERR_LIC_BAD_VERSION = 2022
#: Cannot establish a connection with a license server 
H_ERR_LIC_CANTCONNECT = 2024
#: Session limit exceeded 
H_ERR_LIC_MAXSESSIONS = 2028
#: All licenses in use 
H_ERR_LIC_MAXUSERS = 2029
#: No license server specified for counted license 
H_ERR_LIC_NO_SERVER_IN_FILE = 2030
#: Can not find feature in the license file 
H_ERR_LIC_NOFEATURE = 2031
#: License file does not support a version this new 
H_ERR_LIC_OLDVER = 2033
#: This platform not authorized by license - running on platform not included in PLATFORMS list 
H_ERR_LIC_PLATNOTLIC = 2034
#: License server busy 
H_ERR_LIC_SERVBUSY = 2035
#: Could not find license.dat 
H_ERR_LIC_NOCONFFILE = 2036
#: Invalid license file syntax 
H_ERR_LIC_BADFILE = 2037
#: Cannot connect to a license server 
H_ERR_LIC_NOSERVER = 2038
#: Invalid host 
H_ERR_LIC_NOTTHISHOST = 2041
#: Feature has expired 
H_ERR_LIC_LONGGONE = 2042
#: Invalid date format in license file 
H_ERR_LIC_BADDATE = 2043
#: Invalid returned data from license server 
H_ERR_LIC_BADCOMM = 2044
#: Cannot find SERVER hostname in network database 
H_ERR_LIC_BADHOST = 2045
#: Cannot write data to license server 
H_ERR_LIC_CANTWRITE = 2047
#: License server does not support this version of this feature 
H_ERR_LIC_SERVLONGGONE = 2051
#: Request for more licenses than this feature supports 
H_ERR_LIC_TOOMANY = 2052
#: Cannot find ethernet device 
H_ERR_LIC_CANTFINDETHER = 2055
#: Cannot read license file 
H_ERR_LIC_NOREADLIC = 2056
#: Date too late for binary format 
H_ERR_LIC_DATE_TOOBIG = 2067
#: Server did not respond to message 
H_ERR_LIC_NOSERVRESP = 2069
#: setsockopt() failed 
H_ERR_LIC_SETSOCKFAIL = 2075
#: Message checksum failure 
H_ERR_LIC_BADCHECKSUM = 2076
#: Internal licensing error 
H_ERR_LIC_INTERNAL_ERROR = 2082
#: Server doesn't support this request 
H_ERR_LIC_NOSERVCAP = 2087
#: This feature is available in a different license pool 
H_ERR_LIC_POOL = 2091
#: Dongle not attached, or can't read dongle 
H_ERR_LIC_NODONGLE = 2300
#: Missing dongle driver 
H_ERR_LIC_NODONGLEDRIVER = 2301
#: Timeout 
H_ERR_LIC_TIMEOUT = 2318
#: Invalid license server certificate 
H_ERR_LIC_INVALID_CERTIFICATE = 2321
#: Invalid license server SSL/TLS certificate 
H_ERR_LIC_INVALID_TLS_CERTIFICATE = 2335
#: Invalid activation request received 
H_ERR_LIC_BAD_ACTREQ = 2339
#: Specified operation is not allowed 
H_ERR_LIC_NOT_ALLOWED = 2345
#: Activation error 
H_ERR_LIC_ACTIVATION = 2348
#: No CodeMeter Runtime installed 
H_ERR_LIC_NO_CM_RUNTIME = 2379
#: Installed CodeMeter Runtime is too old 
H_ERR_LIC_CM_RUNTIME_TOO_OLD = 2380
#: License is for wrong HALCON edition 
H_ERR_LIC_WRONG_EDITION = 2381
#: License contains unknown FLAGS 
H_ERR_LIC_UNKNOWN_FLAGS = 2382
#: HALCON preview version expired 
H_ERR_LIC_PREVIEW_EXPIRED = 2383
#: License does not support a HALCON version this old 
H_ERR_LIC_NEWVER = 2384
#: Error codes concerning the HALCON core, 2100..2199 
H_ERR_LIC_RANGE1_BEGIN = H_ERR_LIC_NO_LICENSE
#: Wrong index for output object parameter 
H_ERR_WOOPI = 2100
#: Wrong index for input object parameter
H_ERR_WIOPI = 2101
#: Wrong index for image object 
H_ERR_WOI = 2102
#: Wrong number region/image component 
H_ERR_WRCN = 2103
#: Wrong relation name 
H_ERR_WRRN = 2104
#: Access to undefined gray value component
H_ERR_AUDI = 2105
#: Wrong image width 
H_ERR_WIWI = 2106
#: Wrong image height 
H_ERR_WIHE = 2107
#: Undefined gray value component 
H_ERR_ICUNDEF = 2108
#: Inconsistent data of data base (typing) 
H_ERR_IDBD = 2200
#: Wrong index for input control parameter 
H_ERR_WICPI = 2201
#: Data of data base not defined (internal error) 
H_ERR_DBDU = 2202
#: legacy: Number of operators too big 
H_ERR_PNTL = 2203
#: User extension not properly installed 
H_ERR_UEXTNI = 2205
#: legacy: Number of packages too large 
H_ERR_NPTL = 2206
#: No such package installed 
H_ERR_NSP = 2207
#: incompatible HALCON versions 
H_ERR_ICHV = 2211
#: incompatible operator interface 
H_ERR_ICOI = 2212
#: wrong extension package id 
H_ERR_XPKG_WXID = 2220
#: wrong operator id 
H_ERR_XPKG_WOID = 2221
#: wrong operator information id 
H_ERR_XPKG_WOIID = 2222
#: Wrong Hctuple array type 
H_ERR_CTPL_WTYP = 2400
#: Wrong Hcpar type 
H_ERR_CPAR_WTYP = 2401
#: Wrong Hctuple index 
H_ERR_CTPL_WIDX = 2402
#: Wrong version of file 
H_ERR_WFV = 2403
#: Wrong handle type 
H_ERR_WRONG_HANDLE_TYPE = 2404
#: wrong vector type 
H_ERR_WVTYP = 2410
#: wrong vector dimension 
H_ERR_WVDIM = 2411
#: Wrong (unknown) HALCON handle 
H_ERR_WHDL = 2450
#: Wrong HALCON id, no data available 
H_ERR_WID = 2451
#: HALCON id out of range 
H_ERR_IDOOR = 2452
#: Handle is NULL 
H_ERR_HANDLE_NULL = 2453
#: Handle was cleared 
H_ERR_HANDLE_CLEARED = 2454
#: Handle type does not serialize 
H_ERR_HANDLE_NOSER = 2455
#: Reference cycles of handles found 
H_ERR_HANDLE_CYCLES = 2456
#: Type mismatch: Control expected, found iconic 
H_ERR_WT_CTRL_EXPECTED = 2460
#: Type mismatc: Iconic expected, control found 
H_ERR_WT_ICONIC_EXPECTED = 2461
#: hlibxpi Init function of an extension * that was build with xpi was not * called 
H_ERR_XPI_INIT_NOT_CALLED = 2500
#: hlib didn't find the init function * of the extension it is connecting to * -> old extension without hlibxpi or * the function export failed 
H_ERR_XPI_NO_INIT_FOUND = 2501
#: Unresolved function in hlibxpi 
H_ERR_XPI_UNRES = 2502
#: HALCON extension requires a HALCON * version that is newer than the * connected hlib 
H_ERR_XPI_HLIB_TOO_OLD = 2503
#: the (major) version of the hlibxpi * which is used by the connecting * extension is too small for hlib 
H_ERR_XPI_XPI_TOO_OLD = 2504
#: the major version of the hlibxpi * which is used by the hlib is too * small 
H_ERR_XPI_MAJOR_TOO_SMALL = 2505
#: the minor version of the hlibxpi * which is used by the hlib is too * small 
H_ERR_XPI_MINOR_TOO_SMALL = 2506
#: Wrong major version in symbol struct * (internal: should not happen) 
H_ERR_XPI_INT_WRONG_MAJOR = 2507
#: HLib version could not be detected 
H_ERR_XPI_UNKNOW_HLIB_VER = 2508
#: Wrong hardware information file format 
H_ERR_HW_WFF = 2800
#: Wrong hardware information file version 
H_ERR_HW_WFV = 2801
#: Error while reading the hardware knowledge
H_ERR_HW_RF = 2802
#: Error while writing the hardware knowledge
H_ERR_HW_WF = 2803
#: Tag not found 
H_ERR_HW_TF = 2804
#: No CPU Info 
H_ERR_HW_CPU = 2805
#: No AOP Info 
H_ERR_HW_AOP = 2806
#: No AOP Info for this HALCON variant 
H_ERR_HW_HVAR = 2807
#: No AOP Info for this HALCON architecture 
H_ERR_HW_HARCH = 2808
#: No AOP Info for specified Operator found 
H_ERR_HW_HOP = 2809
#: undefined AOP model 
H_ERR_HW_WAOPM = 2810
#: wrong tag derivate 
H_ERR_HW_WTD = 2811
#: internal error 
H_ERR_HW_IE = 2812
#: hw check was canceled 
H_ERR_HW_CANCEL = 2813
#: Wrong access to global variable 
H_ERR_GV_WA = 2830
#: Used global variable does not exist 
H_ERR_GV_NC = 2831
#: Used global variable not accessible via GLOBAL_ID 
H_ERR_GV_NG = 2832
#: Halcon server to terminate is still working on a job 
H_ERR_HM_NT = 2835
#: No such HALCON software agent 
H_ERR_HM_NA = 2837
#: Hardware check for parallelization not possible on a single-processor machine 
H_ERR_AG_CN = 2838
#: (Seq.) HALCON does not support parallel hardware check (use Parallel HALCON instead) 
H_ERR_AG_NC = 2839
#: Initialization of agent failed 
H_ERR_AG_IN = 2840
#: Termination of agent failed 
H_ERR_AG_NT = 2841
#: Inconsistent hardware description file 
H_ERR_AG_HW = 2842
#: Inconsistent agent information file 
H_ERR_AG_II = 2843
#: Inconsistent agent knowledge file 
H_ERR_AG_IK = 2844
#: The file with the parallelization information does not match to the currently HALCON version/revision 
H_ERR_AG_WV = 2845
#: The file with the parallelization information does not match to the currently used machine 
H_ERR_AG_WH = 2846
#: Inconsistent knowledge base of HALCON software agent 
H_ERR_AG_KC = 2847
#: Unknown communication type 
H_ERR_AG_CT = 2848
#: Unknown message type for HALCON software agent 
H_ERR_AG_MT = 2849
#: Error while saving the parallelization knowledge 
H_ERR_AG_WK = 2850
#: Wrong type of work information 
H_ERR_AG_WW = 2851
#: Wrong type of application information 
H_ERR_AG_WA = 2852
#: Wrong type of experience information 
H_ERR_AG_WE = 2853
#: Unknown name of HALCON software agent 
H_ERR_AG_NU = 2854
#: Unknown name and communication address of HALCON software agent 
H_ERR_AG_NE = 2855
#: cpu representative (HALCON software agent) not reachable 
H_ERR_AG_RR = 2856
#: cpu refuses work 
H_ERR_AG_CR = 2857
#: Description of scheduling resource not found 
H_ERR_AG_RN = 2858
#: Not accessible function of HALCON software agent 
H_ERR_AG_TILT = 2859
#: Wrong type: HALCON scheduling resource 
H_ERR_WRT = 2860
#: Wrong state: HALCON scheduling resource 
H_ERR_WRS = 2861
#: Unknown parameter type: HALCON scheduling resource 
H_ERR_UNKPT = 2862
#: Unknown parameter value: HALCON scheduling resource 
H_ERR_UNKPARVAL = 2863
#: Wrong post processing of control parameter 
H_ERR_CTRL_WPP = 2864
#: Error while trying to get time 
H_ERR_GETTI = 2867
#: Error while trying to get the number of processors 
H_ERR_GETCPUNUM = 2868
#: Error while accessing temporary file 
H_ERR_TMPFNF = 2869
#: message queue wait operation canceled 
H_ERR_MQCNCL = 2890
#: message queue overflow 
H_ERR_MQOVL = 2891
#: Threads still wait on message queue while * clearing it. 
H_ERR_MQCLEAR = 2892
#: Invalid file format for a message 
H_ERR_M_WRFILE = 2893
#: Dict does not contain requested key 
H_ERR_DICT_KEY = 2894
#: Incorrect tuple length in dict 
H_ERR_DICT_TUPLE_LENGTH = 2895
#: Incorrect tuple type in dict 
H_ERR_DICT_TUPLE_TYPE = 2896
#: Invalid index for dict tuple 
H_ERR_DICT_INVALID_INDEX = 2897
#: Error while forcing a context switch 
H_ERR_PTHRD_SCHED = 2900
#: Error while accessing cpu affinity 
H_ERR_SCHED_GAFF = 2901
#: Error while setting cpu affinity 
H_ERR_SCHED_SAFF = 2902
#: wrong synchronization object 
H_ERR_CO_WSO = 2950
#: wrong operator call object 
H_ERR_CO_WOCO = 2952
#: input object not initialized 
H_ERR_CO_IOPNI = 2953
#: input control not initialized 
H_ERR_CO_ICPNI = 2954
#: output object not initialized 
H_ERR_CO_OOPNI = 2955
#: output control not initialized 
H_ERR_CO_OCPNI = 2956
#: Creation of pthread failed 
H_ERR_PTHRD_CR = 2970
#: pthread-detach failed 
H_ERR_PTHRD_DT = 2971
#: pthread-join failed 
H_ERR_PTHRD_JO = 2972
#: Initialization of mutex variable failed 
H_ERR_PTHRD_MI = 2973
#: Deletion of mutex variable failed 
H_ERR_PTHRD_MD = 2974
#: Lock of mutex variable failed 
H_ERR_PTHRD_ML = 2975
#: Unlock of mutex variable failed 
H_ERR_PTHRD_MU = 2976
#: Failed to signal pthread condition var. 
H_ERR_PTHRD_CS = 2977
#: Failed to wait for pthread cond. var. 
H_ERR_PTHRD_CW = 2978
#: Failed to init pthread condition var. 
H_ERR_PTHRD_CI = 2979
#: Failed to destroy pthread condition var.
H_ERR_PTHRD_CD = 2980
#: Failed to signal event. 
H_ERR_PTHRD_ES = 2981
#: Failed to wait for event. 
H_ERR_PTHRD_EW = 2982
#: Failed to init event. 
H_ERR_PTHRD_EI = 2983
#: Failed to destroy event.
H_ERR_PTHRD_ED = 2984
#: Failed to create a tsd key.
H_ERR_PTHRD_TSDC = 2985
#: Failed to set a thread specific data key.
H_ERR_PTHRD_TSDS = 2986
#: Failed to get a tsd key.
H_ERR_PTHRD_TSDG = 2987
#: Failed to free a tsd key.
H_ERR_PTHRD_TSDF = 2988
#: Aborted waiting at a barrier
H_ERR_PTHRD_BA = 2989
#: 'Free list' is empty while scheduling 
H_ERR_DCDG_FLE = 2990
#: Communication partner not checked in 
H_ERR_MSG_PNCI = 2991
#: The communication system can't be started while running 
H_ERR_MSG_CSAI = 2992
#: Communication partner not checked in 
H_ERR_MSG_CSNI = 2993
#: Initialization of Barrier failed 
H_ERR_PTHRD_BI = 2994
#: Waiting at a barrier failed 
H_ERR_PTHRD_BW = 2995
#: Destroying of an barrier failed 
H_ERR_PTHRD_BD = 2996
#: Region completely outside of the image domain 
H_ERR_RCOIMA = 3010
#: Region (partially) outside of the definition range of the image 
H_ERR_ROOIMA = 3011
#: Intersected definition range region/image empty 
H_ERR_RIEI = 3012
#: Image with empty definition range 
H_ERR_EDEF = 3013
#: No common image point of two images 
H_ERR_IIEI = 3014
#: Wrong region for image (first row < 0) 
H_ERR_FLTS = 3015
#: Wrong region for image (column in last row >= image width) 
H_ERR_LLTB = 3016
#: Number of images unequal in input pars. 
H_ERR_UENOI = 3017
#: Image height too small 
H_ERR_HTS = 3018
#: Image width too small 
H_ERR_WTS = 3019
#: Internal error: Multiple call of HRLInitSeg() 
H_ERR_CHSEG = 3020
#: Internal error: HRLSeg() not initialized 
H_ERR_RLSEG1 = 3021
#: Wrong size of filter for Gauss 
H_ERR_WGAUSSM = 3022
#: Filter size exceeds image size 
H_ERR_FSEIS = 3033
#: Filter size evan 
H_ERR_FSEVAN = 3034
#: Filter size to big 
H_ERR_FSTOBIG = 3035
#: Region is empty 
H_ERR_EMPTREG = 3036
#: Domains of the input images differ 
H_ERR_DOM_DIFF = 3037
#: Row value of a coordinate > 2^15-1 (XL: 2^30 - 1) 
H_ERR_ROWTB = 3040
#: Row value of a coordinate < -2^15+1 (XL: -2^30+1) 
H_ERR_ROWTS = 3041
#: Column value of a coordinate > 2^15-1 (XL: 2^30 - 1) 
H_ERR_COLTB = 3042
#: Column value of a coordinate < -2^15+1 (XL: -2^30+1) 
H_ERR_COLTS = 3043
#: Wrong segmentation threshold 
H_ERR_WRTHR = 3100
#: Unknown feature 
H_ERR_UNKF = 3101
#: Unknown gray value feature 
H_ERR_UNKG = 3102
#: Internal error in HContCut 
H_ERR_EINCC = 3103
#: Error in HContToPol: distance of points too big 
H_ERR_EINCP1 = 3104
#: Error in HContToPol: contour too long 
H_ERR_EINCP2 = 3105
#: Too many rows (IPImageTransform) 
H_ERR_TMR = 3106
#: Scaling factor = 0.0 (IPImageScale) 
H_ERR_SFZ = 3107
#: Wrong range in transformation matrix 
H_ERR_OOR = 3108
#: Internal error in IPvvf: no element free 
H_ERR_NEF = 3109
#: Number of input objects is zero 
H_ERR_NOOB = 3110
#: At least one input object has an empty region 
H_ERR_EMPOB = 3111
#: Operation allowed for rectangular images 2**n only 
H_ERR_NPOT = 3112
#: Too many relevant points (IPHysterese) 
H_ERR_TMEP = 3113
#: Number of labels in image too big 
H_ERR_LTB = 3114
#: No labels with negative values allowed 
H_ERR_NNLA = 3115
#: Wrong filter size (too small ?) 
H_ERR_WFS = 3116
#: Images with different image size 
H_ERR_IWDS = 3117
#: Target image too wide or too far on the right 
H_ERR_IWTL = 3118
#: Target image too narrow or too far on the left 
H_ERR_IWTS = 3119
#: Target image too high or too far down 
H_ERR_IHTL = 3120
#: Target image too low or too far up 
H_ERR_IHTS = 3121
#: Number of channels in the input parameters are different 
H_ERR_DNOC = 3122
#: Wrong color filter array type 
H_ERR_WRCFAFLT = 3123
#: Wrong color filter array interpolation 
H_ERR_WRCFAINT = 3124
#: Homogeneous matrix does not represent an affine transformation 
H_ERR_NO_AFFTRANS = 3125
#: Inpainting region too close to the image border 
H_ERR_INPNOBDRY = 3126
#: source and destination differ in size
H_ERR_DSIZESD = 3127
#: to many features
H_ERR_TMFEAT = 3128
#: Reflection axis undefined 
H_ERR_AXIS_UNDEF = 3129
#: Coocurrence Matrix: Too little columns for quantisation 
H_ERR_COWTS = 3131
#: Coocurrence Matrix: Too little rows for quantisation 
H_ERR_COHTS = 3132
#: Wrong number of columns 
H_ERR_NUM_COLMN = 3133
#: Wrong number of rows 
H_ERR_NUM_LINES = 3134
#: Number has too many digits 
H_ERR_OVL = 3135
#: Matrix is not symmetric 
H_ERR_NOT_SYM = 3136
#: Matrix is too big 
H_ERR_NUM_COLS = 3137
#: Wrong structure of file 
H_ERR_SYNTAX = 3138
#: Less than 2 matrices 
H_ERR_MISSING = 3139
#: Not enough memory 
H_ERR_COOC_MEM = 3140
#: Can not read the file 
H_ERR_NO_FILE = 3141
#: Can not open file for writing 
H_ERR_FILE_WR = 3142
#: Too many lookup table colors 
H_ERR_NUM_LUCOLS = 3143
#: Too many Hough points (lines) 
H_ERR_WNOLI = 3145
#: Target image has got wrong height (not big enough) 
H_ERR_DITS = 3146
#: Wrong interpolation mode 
H_ERR_WINTM = 3147
#: Region not compact or not connected 
H_ERR_THICK_NK = 3148
#: Wrong filter index for filter size 3 
H_ERR_WIND3 = 3170
#: Wrong filter index for filter size 5 
H_ERR_WIND5 = 3171
#: Wrong filter index for filter size 7 
H_ERR_WIND7 = 3172
#: Wrong filter size; only 3/5/7 
H_ERR_WLAWSS = 3173
#: Number of suitable pixels too small to reliably estimate the noise 
H_ERR_NE_NPTS = 3175
#: Different number of entries/exits in HContCut 
H_ERR_WNEE = 3200
#: Reference to contour is missing 
H_ERR_REF = 3201
#: Wrong XLD type 
H_ERR_XLDWT = 3250
#: Border point is set to FG 
H_ERR_XLD_RPF = 3252
#: Maximum contour length exceeded 
H_ERR_XLD_MCL = 3253
#: Maximum number of contours exceeded 
H_ERR_XLD_MCN = 3254
#: Contour too short for fetch_angle_xld 
H_ERR_XLD_CTS = 3255
#: Regression parameters of contours already computed 
H_ERR_XLD_CRD = 3256
#: Regression parameters of contours not yet entered! 
H_ERR_XLD_CRND = 3257
#: Data base: XLD object has been deleted 
H_ERR_DBXC = 3258
#: Data base: Object has no XLD-ID 
H_ERR_DBWXID = 3259
#: Wrong number of contour points allocated 
H_ERR_XLD_WNP = 3260
#: Contour attribute not defined 
H_ERR_XLD_CAND = 3261
#: Ellipse fitting failed 
H_ERR_FIT_ELLIPSE = 3262
#: Circle fitting failed 
H_ERR_FIT_CIRCLE = 3263
#: All points classified as outliers (ClippingFactor too small or used points not similar to primitive) 
H_ERR_FIT_CLIP = 3264
#: Quadrangle fitting failed 
H_ERR_FIT_QUADRANGLE = 3265
#: No points for at least one rectangle side 
H_ERR_INCOMPL_RECT = 3266
#: A contour point lies outside of the image 
H_ERR_XLD_COI = 3267
#: Not enough points for model fitting 
H_ERR_FIT_NOT_ENOUGH_POINTS = 3274
#: No ARC/INFO world file 
H_ERR_NWF = 3275
#: No ARC/INFO generate file 
H_ERR_NAIGF = 3276
#: Unexpected end of file while reading DXF file 
H_ERR_DXF_UEOF = 3278
#: Cannot read DXF-group code from file 
H_ERR_DXF_CRGC = 3279
#: Inconsistent number of attributes per point in DXF file 
H_ERR_DXF_INAPP = 3280
#: Inconsistent number of attributes and names in DXF file 
H_ERR_DXF_INAPPN = 3281
#: Inconsistent number of global attributes and names in DXF file 
H_ERR_DXF_INAPCN = 3282
#: Cannot read attributes from DXF file 
H_ERR_DXF_CRAPP = 3283
#: Cannot read global attributes from DXF file 
H_ERR_DXF_CRAPC = 3284
#: Cannot read attribute names from DXF file 
H_ERR_DXF_CRAN = 3285
#: Wrong generic parameter name 
H_ERR_DXF_WPN = 3286
#: Internal DXF I/O error: Wrong data type 
H_ERR_DXF_IEDT = 3289
#: Isolated point while contour merging 
H_ERR_XLD_ISOL_POINT = 3290
#: Constraints cannot be fulfilled 
H_ERR_NURBS_CCBF = 3291
#: No segment in contour 
H_ERR_NSEG = 3292
#: Only one or no point in template contour 
H_ERR_NO_ONE_P = 3293
#: Syntax error in file for training 
H_ERR_SESF = 3300
#: Maximum number of attributes per example exceeded 
H_ERR_TMFE = 3301
#: Not possible to open file for training 
H_ERR_OPSF = 3302
#: Too many data sets for training 
H_ERR_TMSS = 3303
#: Too many examples for one data set for training 
H_ERR_TMSAM = 3305
#: Too many classes 
H_ERR_TMCLS = 3306
#: Maximum number of cuboids exceeded 
H_ERR_TMBOX = 3307
#: Not possible to open classificator's file 
H_ERR_OPCF = 3308
#: Error while saving the classificator 
H_ERR_SCLA = 3309
#: Not possible to open protocol file 
H_ERR_OPF = 3310
#: Classificator with this name is already existent 
H_ERR_CLEX = 3311
#: Maximum number of classificators exceeded 
H_ERR_TMCLA = 3312
#: Name of classificator is too long, >= 20 
H_ERR_CNTL = 3313
#: Classificator with this name is not existent 
H_ERR_CLNNF = 3314
#: Current classificator is not defined 
H_ERR_NCCLA = 3315
#: Wrong id in classification file 
H_ERR_CLASS2_ID = 3316
#: The version of the classifier is not supported 
H_ERR_CLASS2_VERS = 3317
#: Serialized item does not contain a valid classifier 
H_ERR_CLASS_NOSITEM = 3318
#: Text model does not contain a classifier yet (use set_text_model_param) 
H_ERR_TM_NO_CL = 3319
#: Adding new features is not possible, because the dataset has been normalized during training. Please create a new classfier and add all training data again or disable normalization during training.
H_ERR_KNN_CANNOT_ADD = 3320
#: Error in KMeans cluter initialization. 
H_ERR_ML_KMEAN_INITIALIZATION_ERROR = 3325
#: Invalid file format for GMM training samples 
H_ERR_GMM_NOTRAINFILE = 3330
#: The version of the GMM training samples is not supported 
H_ERR_GMM_WRTRAINVERS = 3331
#: Wrong training sample file format 
H_ERR_GMM_WRSMPFORMAT = 3332
#: nvalid file format for Gaussian Mixture Model (GMM) 
H_ERR_GMM_NOCLASSFILE = 3333
#: The version of the Gaussian Mixture Model (GMM) is not supported 
H_ERR_GMM_WRCLASSVERS = 3334
#: Unknown error when training GMM 
H_ERR_GMM_TRAIN_UNKERR = 3335
#: Collapsed covariance matrix 
H_ERR_GMM_TRAIN_COLLAPSED = 3336
#: No samples for at least one class 
H_ERR_GMM_TRAIN_NOSAMPLE = 3337
#: Too few samples for at least one class 
H_ERR_GMM_TRAIN_FEWSAMPLES = 3338
#: GMM is not trained 
H_ERR_GMM_NOTTRAINED = 3340
#: GMM has no training data 
H_ERR_GMM_NOTRAINDATA = 3341
#: Serialized item does not contain a valid Gaussian Mixture Model (GMM) 
H_ERR_GMM_NOSITEM = 3342
#: Unknown output function 
H_ERR_MLP_UNKOUTFUNC = 3350
#: Target not in 0-1 encoding 
H_ERR_MLP_NOT01ENC = 3351
#: No training samples stored in the classifier 
H_ERR_MLP_NOTRAINDATA = 3352
#: Invalid file format for MLP training samples 
H_ERR_MLP_NOTRAINFILE = 3353
#: The version of the MLP training samples is not supported 
H_ERR_MLP_WRTRAINVERS = 3354
#: Wrong training sample format 
H_ERR_MLP_WRSMPFORMAT = 3355
#: MLP is not a classifier 
H_ERR_MLP_NOCLASSIF = 3356
#: Invalid file format for multilayer perceptron (MLP) 
H_ERR_MLP_NOCLASSFILE = 3357
#: The version of the multilayer perceptron (MLP) is not supported 
H_ERR_MLP_WRCLASSVERS = 3358
#: Wrong number of channels 
H_ERR_WRNUMCHAN = 3359
#: Wrong number of MLP parameters 
H_ERR_MLP_WRNUMPARAM = 3360
#: Serialized item does not contain a valid multilayer perceptron (MLP) 
H_ERR_MLP_NOSITEM = 3361
#: The number of image channels and the number of dimensions of the look-up table do not match 
H_ERR_LUT_WRNUMCHAN = 3370
#: A look-up table can be build for 2 or 3 channels only 
H_ERR_LUT_NRCHANLARGE = 3371
#: Cannot create look-up table. Please choose a larger 'bit_depth' or select the 'fast' 'class_selection'. 
H_ERR_LUT_CANNOTCREAT = 3372
#: No training samples stored in the classifier 
H_ERR_SVM_NOTRAINDATA = 3380
#: Invalid file format for SVM training samples 
H_ERR_SVM_NOTRAINFILE = 3381
#: The version of the SVM training samples is not supported 
H_ERR_SVM_WRTRAINVERS = 3382
#: Wrong training sample format 
H_ERR_SVM_WRSMPFORMAT = 3383
#: Invalid file format for support vector machine (SVM) 
H_ERR_SVM_NOCLASSFILE = 3384
#: The version of the support vector machine (SVM) is not supported 
H_ERR_SVM_WRCLASSVERS = 3385
#: Wrong number of classes 
H_ERR_SVM_WRNRCLASS = 3386
#: Chosen nu is too big 
H_ERR_SVM_NU_TOO_BIG = 3387
#: SVM Training failed 
H_ERR_SVM_TRAIN_FAIL = 3388
#: SVMs do not fit together 
H_ERR_SVM_DO_NOT_FIT = 3389
#: No SV in SVM to add to training 
H_ERR_SVM_NO_TRAIN_ADD = 3390
#: Kernel must be RBF 
H_ERR_SVM_KERNELNOTRBF = 3391
#: Not all classes contained in training data 
H_ERR_SVM_NO_TRAIND_FOR_CLASS = 3392
#: SVM not trained 
H_ERR_SVM_NOT_TRAINED = 3393
#: Classifier not trained 
H_ERR_NOT_TRAINED = 3394
#: Serialized item does not contain a valid support vector machine (SVM) 
H_ERR_SVM_NOSITEM = 3395
#: Wrong rotation number 
H_ERR_ROTNR = 3401
#: Wrong letter for Golay element 
H_ERR_GOL = 3402
#: Wrong reference point 
H_ERR_BEZ = 3403
#: Wrong number of iterations 
H_ERR_ITER = 3404
#: Mophology: system error 
H_ERR_MOSYS = 3405
#: Wrong type of boundary 
H_ERR_ART = 3406
#: Morphology: Wrong number of input obj. 
H_ERR_OBJI = 3407
#: Morphology: Wrong number of output obj. 
H_ERR_OBJO = 3408
#: Morphology: Wrong number of input control parameter 
H_ERR_PARI = 3409
#: Morphology: Wrong number of output control parameter 
H_ERR_PARO = 3410
#: Morphology: Struct. element is infinite 
H_ERR_SELC = 3411
#: Morphology: Wrong name for struct. elem. 
H_ERR_WRNSE = 3412
#: Wrong number of run length rows (chords): smaller than 0 
H_ERR_WRRLN1 = 3500
#: Number of chords too big, increase * current_runlength_number using set_system
H_ERR_WRRLN2 = 3501
#: Run length row with negative length 
H_ERR_WRRLL = 3502
#: Run length row >= image height 
H_ERR_RLLTB = 3503
#: Run length row < 0 
H_ERR_RLLTS = 3504
#: Run length column >= image width 
H_ERR_RLCTB = 3505
#: Lauflaengenspalte < 0 
H_ERR_RLCTS = 3506
#: For CHORD_TYPE: Number of row too big 
H_ERR_CHLTB = 3507
#: For CHORD_TYPE: Number of row too small 
H_ERR_CHLTS = 3508
#: For CHORD_TYPE: Number of column too big 
H_ERR_CHCTB = 3509
#: Exceeding the maximum number of run lengths while automatic expansion 
H_ERR_MRLE = 3510
#: Region->compl neither true/false 
H_ERR_ICCOMPL = 3511
#: Region->max_num < Region->num 
H_ERR_RLEMAX = 3512
#: Number of chords too big for num_max 
H_ERR_WRRLN3 = 3513
#: Operator cannot be implemented for complemented regions 
H_ERR_OPNOCOMPL = 3514
#: Image width < 0 
H_ERR_WIMAW1 = 3520
#: Image width >= MAX_FORMAT 
H_ERR_WIMAW2 = 3521
#: Image height <= 0 
H_ERR_WIMAH1 = 3522
#: Image height >= MAX_FORMAT 
H_ERR_WIMAH2 = 3523
#: Image width <= 0 
H_ERR_WIMAW3 = 3524
#: Image height <= 0 
H_ERR_WIMAH3 = 3525
#: Too many segments 
H_ERR_TMS = 3550
#: INT8 images are available on 64 bit systems only 
H_ERR_NO_INT8_IMAGE = 3551
#: Point at infinity cannot be converted to a Euclidean point 
H_ERR_POINT_AT_INFINITY = 3600
#: Covariance matrix could not be determined 
H_ERR_ML_NO_COVARIANCE = 3601
#: RANSAC algorithm didn't find enough point correspondences 
H_ERR_RANSAC_PRNG = 3602
#: RANSAC algorithm didn't find enough point correspondences 
H_ERR_RANSAC_TOO_DIFFERENT = 3603
#: Internal diagnosis: fallback method had to be used 
H_ERR_PTI_FALLBACK = 3604
#: Projective transformation is singular 
H_ERR_PTI_TRAFO_SING = 3605
#: Mosaic is under-determined 
H_ERR_PTI_MOSAIC_UNDERDET = 3606
#: Input covariance matrix is not positive definite 
H_ERR_COV_NPD = 3607
#: The number of input points too large. 
H_ERR_TOO_MANY_POINTS = 3608
#: Inconsistent number of point correspondences. 
H_ERR_INPC = 3620
#: No path from reference image to one or more images. 
H_ERR_NOPA = 3621
#: Image with specified index does not exist. 
H_ERR_IINE = 3622
#: Matrix is not a camera matrix. 
H_ERR_NOCM = 3623
#: Skew is not zero. 
H_ERR_SKNZ = 3624
#: Illegal focal length. 
H_ERR_ILFL = 3625
#: Kappa is not zero. 
H_ERR_KANZ = 3626
#: It is not possible to determine all parameters for in the variable case. 
H_ERR_VARA = 3627
#: No valid implementation selected. 
H_ERR_LVDE = 3628
#: Kappa can only be determined with the gold-standard method for fixed camera parameters. 
H_ERR_KPAR = 3629
#: Conflicting number of images and projection mode. 
H_ERR_IMOD = 3630
#: Error in projection: Point not in any cube map. 
H_ERR_PNIC = 3631
#: No solution found. 
H_ERR_NO_SOL = 3632
#: Tilt is not zero. 
H_ERR_TINZ = 3633
#: Illegal combination of parameters and estimation method. 
H_ERR_ILMD = 3640
#: No suitable contours found 
H_ERR_RDS_NSC = 3660
#: No stable solution found 
H_ERR_RDS_NSS = 3661
#: Instable solution found 
H_ERR_RDS_ISS = 3662
#: Not enough contours for calibration 
H_ERR_RDS_NEC = 3663
#: Invalid file format for FFT optimization data 
H_ERR_NOFFTOPT = 3650
#: The version of the FFT optimization data is not supported 
H_ERR_WRFFTOPTVERS = 3651
#: Optimization data was created with a different HALCON version (Standard HALCON / Parallel HALCON) 
H_ERR_WRHALCONVERS = 3652
#: Storing of the optimization data failed 
H_ERR_OPTFAIL = 3653
#: Serialized item does not contain valid FFT optimization data 
H_ERR_FFTOPT_NOSITEM = 3654
#: Invalid disparity range for binocular_disparity_ms method 
H_ERR_INVLD_DISP_RANGE = 3690
#: Epipoles are situated within the image domain 
H_ERR_EPIINIM = 3700
#: Fields of view of both cameras do not intersect each other 
H_ERR_EPI_FOV = 3701
#: Rectification impossible 
H_ERR_EPI_RECT = 3702
#: Wrong type of target_thickness parameter 
H_ERR_BI_WT_TARGET = 3710
#: Wrong type of thickness_tolerance parameter 
H_ERR_BI_WT_THICKNESS = 3711
#: Wrong type of position_tolerance parameter 
H_ERR_BI_WT_POSITION = 3712
#: Wrong type of sigma parameter 
H_ERR_BI_WT_SIGMA = 3713
#: Wrong value of sigma parameter 
H_ERR_BI_WV_SIGMA = 3714
#: Wrong type of threshold parameter 
H_ERR_BI_WT_THRESH = 3715
#: Wrong value of target_thickness parameter 
H_ERR_BI_WV_TARGET = 3716
#: Wrong value of thickness_tolerance parameter 
H_ERR_BI_WV_THICKNESS = 3717
#: Wrong value of position_tolerance parameter 
H_ERR_BI_WV_POSITION = 3718
#: Wrong value of threshold parameter 
H_ERR_BI_WV_THRESH = 3719
#: Wrong type of refinement parameter 
H_ERR_BI_WT_REFINE = 3720
#: Wrong value of refinement parameter 
H_ERR_BI_WV_REFINE = 3721
#: Wrong type of resolution parameter 
H_ERR_BI_WT_RESOL = 3722
#: Wrong type of resolution parameter 
H_ERR_BI_WV_RESOL = 3723
#: Wrong type of polarity parameter 
H_ERR_BI_WT_POLARITY = 3724
#: Wrong type of polarity parameter 
H_ERR_BI_WV_POLARITY = 3725
#: No sheet-of-light model available
H_ERR_SOL_EMPTY_MODEL_LIST = 3751
#: Wrong input image size (width) 
H_ERR_SOL_WNIW = 3752
#: Wrong input image size (height) 
H_ERR_SOL_WNIH = 3753
#: profile region does not fit the domain of definition of the input image 
H_ERR_SOL_WPROF_REG = 3754
#: Calibration extend not set 
H_ERR_SOL_CAL_NONE = 3755
#: Undefined disparity image 
H_ERR_SOL_UNDEF_DISPARITY = 3756
#: Undefined domain for disparity image 
H_ERR_SOL_UNDEF_DISPDOMAIN = 3757
#: Undefined camera parameter 
H_ERR_SOL_UNDEF_CAMPAR = 3758
#: Undefined pose of the lightplane 
H_ERR_SOL_UNDEF_LPCS = 3759
#: Undefined pose of the camera coordinate system 
H_ERR_SOL_UNDEF_CCS = 3760
#: Undefined transformation from the camera to the lightplane coordinate system 
H_ERR_SOL_UNDEF_CCS_2_LPCS = 3761
#: Undefined movement pose for xyz calibration 
H_ERR_SOL_UNDEF_MOV_POSE = 3762
#: Wrong value of scale parameter 
H_ERR_SOL_WV_SCALE = 3763
#: Wrong parameter name 
H_ERR_SOL_WV_PAR_NAME = 3764
#: Wrong type of parameter method 
H_ERR_SOL_WT_METHOD = 3765
#: Wrong type of parameter ambiguity 
H_ERR_SOL_WT_AMBIGUITY = 3766
#: Wrong type of parameter score 
H_ERR_SOL_WT_SCORE_TYPE = 3767
#: Wrong type of parameter calibration 
H_ERR_SOL_WT_CALIBRATION = 3768
#: Wrong type of parameter number_profiles 
H_ERR_SOL_WT_NUM_PROF = 3769
#: Wrong type of element in parameter camera_parameter 
H_ERR_SOL_WT_CAM_PAR = 3770
#: Wrong type of element in pose 
H_ERR_SOL_WT_PAR_POSE = 3771
#: Wrong value of parameter method 
H_ERR_SOL_WV_METHOD = 3772
#: Wrong type of parameter min_gray 
H_ERR_SOL_WT_THRES = 3773
#: Wrong value of parameter ambiguity 
H_ERR_SOL_WV_AMBIGUITY = 3774
#: Wrong value of parameter score_type 
H_ERR_SOL_WV_SCORE_TYPE = 3775
#: Wrong value of parameter calibration 
H_ERR_SOL_WV_CALIBRATION = 3776
#: Wrong value of parameter number_profiles 
H_ERR_SOL_WV_NUM_PROF = 3777
#: Wrong type of camera 
H_ERR_SOL_WV_CAMERA_TYPE = 3778
#: Wrong number of values of parameter camera_parameter 
H_ERR_SOL_WN_CAM_PAR = 3779
#: Wrong number of values of pose 
H_ERR_SOL_WN_POSE = 3780
#: Calibration target not found 
H_ERR_SOL_NO_TARGET_FOUND = 3781
#: The calibration algorithm failed to find a valid solution. 
H_ERR_SOL_NO_VALID_SOL = 3782
#: Wrong type of parameter calibration_object 
H_ERR_SOL_WT_CALIB_OBJECT = 3783
#: Invalid calibration object 
H_ERR_SOL_INVALID_CALIB_OBJECT = 3784
#: No calibration object set 
H_ERR_SOL_NO_CALIB_OBJECT_SET = 3785
#: Invalid file format for sheet-of-light model 
H_ERR_SOL_WR_FILE_FORMAT = 3786
#: The version of the sheet-of-light model is not supported 
H_ERR_SOL_WR_FILE_VERS = 3787
#: Camera type not supported by calibrate_sheet_of_light_model
H_ERR_SOL_CAMPAR_UNSUPPORTED = 3788
#: Parameter does not match the set 'calibration' 
H_ERR_SOL_PAR_CALIB = 3790
#: The gray values of the disparity image do not fit the height of the camera 
H_ERR_SOL_WGV_DISP = 3791
#: Wrong texture inspection model type
H_ERR_TI_WRONGMODEL = 3800
#: Texture Model is not trained 
H_ERR_TI_NOTTRAINED = 3801
#: Texture Model has no training data 
H_ERR_TI_NOTRAINDATA = 3802
#: Invalid file format for Texture inspection model 
H_ERR_TI_NOTRAINFILE = 3803
#: The version of the Texture inspection model is not supported 
H_ERR_TI_WRTRAINVERS = 3804
#: Wrong training sample file format 
H_ERR_TI_WRSMPFORMAT = 3805
#: The version of the training sample file is not supported 
H_ERR_TI_WRSMPVERS = 3806
#: At least one of the images is too small 
H_ERR_TI_WRIMGSIZE = 3807
#: The samples do not match the current texture model 
H_ERR_TI_WRSMPTEXMODEL = 3808
#: No images within the texture model 
H_ERR_NOT_ENOUGH_IMAGES = 3809
#: The light source positions are linearly dependent 
H_ERR_SING = 3850
#: No sufficient image indication 
H_ERR_FEWIM = 3851
#: Internal error: Function has equal signs in HZBrent 
H_ERR_ZBR_NOS = 3852
#: Kalman: Dimension n,m or p has got a undefined value 
H_ERR_DIMK = 3900
#: Kalman: File does not exist 
H_ERR_NOFILE = 3901
#: Kalman: Error in file (row of dimension) 
H_ERR_FF1 = 3902
#: Kalman: Error in file (row of marking) 
H_ERR_FF2 = 3903
#: Error in file (value is no float) 
H_ERR_FF3 = 3904
#: Kalman: Matrix A is missing in file 
H_ERR_NO_A = 3905
#: Kalman: In Datei fehlt Matrix C 
H_ERR_NO_C = 3906
#: Kalman: Matrix Q is missing in file 
H_ERR_NO_Q = 3907
#: Kalman: Matrix R is missing in file 
H_ERR_NO_R = 3908
#: Kalman: G or u is missing in file 
H_ERR_NO_GU = 3909
#: Kalman: Covariant matrix is not symmetric 
H_ERR_NOTSYMM = 3910
#: Kalman: Equation system is singular 
H_ERR_SINGU = 3911
#: structured light model is not in persistent mode 
H_ERR_SLM_NOT_PERSISTENT = 3950
#: the min_stripe_width is too large for the chosen pattern_width or pattern_height 
H_ERR_SLM_MSW_TOO_LARGE = 3951
#: the single_stripe_width is too large for the chosen pattern_width or pattern_height 
H_ERR_SLM_SSW_TOO_LARGE = 3952
#: min_stripe_width has to be smaller than single_stripe_width. 
H_ERR_SLM_MSW_GT_SSW = 3953
#: single_stripe_width is too small for min_stripe_width. 
H_ERR_SLM_SSW_LT_MSW = 3954
#: The SLM is not prepared for decoding. 
H_ERR_SLM_NOT_PREP = 3955
#: The SLM does not contain the queried object. 
H_ERR_SLM_NO_OBJS = 3956
#: The version of the structured light model is not supported 
H_ERR_SLM_WRVERS = 3957
#: Invalid file format for a structured light model 
H_ERR_SLM_WRFILE = 3958
#: Wrong pattern type
H_ERR_SLM_WRONGPATTERN = 3959
#: The SLM is not decoded for deflectometry. 
H_ERR_SLM_NOT_DECODED = 3960
#: Wrong model type
H_ERR_SLM_WRONGMODEL = 3961
#: The csm has to contain two camera parameters 
H_ERR_SLM_WNUMCAMS = 3962
#: Inconsistent projector size 
H_ERR_SLM_WPATTSIZE = 3963
#: Camera type not supported 
H_ERR_SLM_WRONGCTYPE = 3964
#: Projector type not supported 
H_ERR_SLM_WRONGPTYPE = 3965
#: The SLM does not contain a csm 
H_ERR_SLM_NO_CSM = 3966
#: The SLM is not set for vertical decoding 
H_ERR_SLM_NO_VERT = 3967
#: The SLM is not decoded for reconstruction 
H_ERR_SLM_NOT_DEC_REC = 3968
#: Inconsistent camera size 
H_ERR_SLM_WCAMSIZE = 3969
#: Object is an object tuple 
H_ERR_DBOIT = 4050
#: Object has been deleted already 
H_ERR_DBOC = 4051
#: Wrong object-ID 
H_ERR_DBWOID = 4052
#: Object tuple has been deleted already 
H_ERR_DBTC = 4053
#: Wrong object tupel-ID 
H_ERR_DBWTID = 4054
#: Object tuple is an object 
H_ERR_DBTIO = 4055
#: Object-ID is NULL (0) 
H_ERR_DBIDNULL = 4056
#: Object-ID outside the valid range 
H_ERR_WDBID = 4057
#: Access to deleted image 
H_ERR_DBIC = 4058
#: Access to image with wrong key 
H_ERR_DBWIID = 4059
#: Access to deleted region 
H_ERR_DBRC = 4060
#: Access to region with wrong key 
H_ERR_DBWRID = 4061
#: Wrong value for image channel 
H_ERR_WCHAN = 4062
#: Index too big 
H_ERR_DBITL = 4063
#: Index not defined 
H_ERR_DBIUNDEF = 4064
#: No OpenCL available 
H_ERR_NO_OPENCL = 4100
#: OpenCL Error occurred 
H_ERR_OPENCL_ERROR = 4101
#: No compute devices available 
H_ERR_NO_COMPUTE_DEVICES = 4102
#: No device implementation for this parameter 
H_ERR_NO_DEVICE_IMPL = 4103
#: Out of device memory 
H_ERR_OUT_OF_DEVICE_MEM = 4104
#: Invalid work group shape 
H_ERR_INVALID_SHAPE = 4105
#: Invalid compute device 
H_ERR_INVALID_DEVICE = 4106
#: CUDA error occurred 
H_ERR_CUDA_ERROR = 4200
#: cuDNN error occurred 
H_ERR_CUDNN_ERROR = 4201
#: cuBLAS error occurred 
H_ERR_CUBLAS_ERROR = 4202
#: Set batch_size not supported 
H_ERR_BATCH_SIZE_NOT_SUPPORTED = 4203
#: CUDA implementations not available 
H_ERR_CUDA_NOT_AVAILABLE = 4204
#: Unsupported version of cuDNN 
H_ERR_CUDNN_UNSUPPORTED_VERSION = 4205
#: Requested feature not supported by cuDNN 
H_ERR_CUDNN_FEATURE_NOT_SUPPORTED = 4206
#: CUDA driver is out-of-date 
H_ERR_CUDA_DRIVER_VERSION = 4207
#: Training is unsupported with the selected runtime. 
H_ERR_TRAINING_UNSUPPORTED = 4301
#: CPU based inference is not supported on this platform 
H_ERR_CPU_INFERENCE_NOT_AVAILABLE = 4302
#: Error occurred in DNNL library 
H_ERR_DNNL_ERROR = 4303
#: AI Accelerator Interface error occurred 
H_ERR_HAI2_ERROR = 4320
#: Invalid parameter for AI Accelerator Interface 
H_ERR_HAI2_INVALID_PARAM = 4321
#: ACL error occurred 
H_ERR_ACL_ERROR = 4400
#: Internal visualization error 
H_ERR_VISUALIZATION = 4500
#: Unexpected color type 
H_ERR_COLOR_TYPE_UNEXP = 4501
#: Number of color settings exceeded 
H_ERR_NUM_COLOR_EXCEEDED = 4502
#: Wrong (logical) window number 
H_ERR_WSCN = 5100
#: Error while opening the window 
H_ERR_DSCO = 5101
#: Wrong window coordinates 
H_ERR_WWC = 5102
#: It is not possible to open another window 
H_ERR_NWA = 5103
#: Device resp. operator not available 
H_ERR_DNA = 5104
#: Unknown color 
H_ERR_UCOL = 5105
#: No window has been opened for desired action 
H_ERR_NWO = 5106
#: Wrong filling mode for regions 
H_ERR_WFM = 5107
#: Wrong gray value (0..255) 
H_ERR_WGV = 5108
#: Wrong pixel value 
H_ERR_WPV = 5109
#: Wrong line width 
H_ERR_WLW = 5110
#: Wrong name of cursor 
H_ERR_WCUR = 5111
#: Wrong color table 
H_ERR_WLUT = 5112
#: Wrong representation mode 
H_ERR_WDM = 5113
#: Wrong representation color 
H_ERR_WRCO = 5114
#: Wrong dither matrix 
H_ERR_WRDM = 5115
#: Wrong image transformation 
H_ERR_WRIT = 5116
#: Unsuitable image type for image trafo. 
H_ERR_IPIT = 5117
#: Wrong zooming factor for image trafo. 
H_ERR_WRZS = 5118
#: Wrong representation mode 
H_ERR_WRDS = 5119
#: Wrong code of device 
H_ERR_WRDV = 5120
#: Wrong number for father window 
H_ERR_WWINF = 5121
#: Wrong window size 
H_ERR_WDEXT = 5122
#: Wrong window type 
H_ERR_WWT = 5123
#: No current window has been set 
H_ERR_WND = 5124
#: Wrong color combination or range (RGB) 
H_ERR_WRGB = 5125
#: Wrong number of pixels set 
H_ERR_WPNS = 5126
#: Wrong value for comprise 
H_ERR_WCM = 5127
#: set_fix with 1/4 image levels and static not valid 
H_ERR_FNA = 5128
#: set_lut not valid in child windows 
H_ERR_LNFS = 5129
#: Number of concurrent used color tables is too big 
H_ERR_LOFL = 5130
#: Wrong device for window dump 
H_ERR_WIDT = 5131
#: Wrong window size for window dump 
H_ERR_WWDS = 5132
#: System variable DISPLAY not defined 
H_ERR_NDVS = 5133
#: Wrong thickness for window margin 
H_ERR_WBW = 5134
#: System variable DISPLAY has been set wrong (<host>:0.0) 
H_ERR_WDVS = 5135
#: Too many fonts loaded 
H_ERR_TMF = 5136
#: Wrong font name 
H_ERR_WFN = 5137
#: No valid cursor position 
H_ERR_WCP = 5138
#: Window is not a textual window 
H_ERR_NTW = 5139
#: Window is not a image window 
H_ERR_NPW = 5140
#: String too long or too high 
H_ERR_STL = 5141
#: Too little space in the window rightw. 
H_ERR_NSS = 5142
#: Window is not suitable for the mouse 
H_ERR_NMS = 5143
#: Here Windows on a equal machine is permitted only 
H_ERR_DWNA = 5144
#: Wrong mode while opening a window 
H_ERR_WOM = 5145
#: Wrong window mode for operation 
H_ERR_WWM = 5146
#: Operation not possible with fixed pixel 
H_ERR_LUTF = 5147
#: Color tables for 8 image levels only 
H_ERR_LUTN8 = 5148
#: Wrong mode for pseudo real colors 
H_ERR_WTCM = 5149
#: Wrong pixel value for LUT 
H_ERR_WIFTL = 5150
#: Wrong image size for pseudo real colors 
H_ERR_WSOI = 5151
#: Error in procedure HRLUT 
H_ERR_HRLUT = 5152
#: Wrong number of entries in color table for set_lut 
H_ERR_WPFSL = 5153
#: Wrong values for image area 
H_ERR_WPVS = 5154
#: Wrong line pattern 
H_ERR_WLPN = 5155
#: Wrong number of parameters for line pattern 
H_ERR_WLPL = 5156
#: Wrong number of colors 
H_ERR_WNOC = 5157
#: Wrong value for mode of area creation 
H_ERR_WPST = 5158
#: Spy window is not set (set_spy) 
H_ERR_SWNA = 5159
#: No file for spy has been set (set_spy) 
H_ERR_NSFO = 5160
#: Wrong parameter output depth (set_spy) 
H_ERR_WSPN = 5161
#: Wrong window size for window dump 
H_ERR_WIFFD = 5162
#: Wrong color table: Wrong file name or query_lut() 
H_ERR_WLUTF = 5163
#: Wrong color table: Empty string ? 
H_ERR_WLUTE = 5164
#: Using this hardware set_lut('default') is allowed only 
H_ERR_WLUTD = 5165
#: Error while calling online help 
H_ERR_CNDP = 5166
#: Row can not be projected 
H_ERR_LNPR = 5167
#: Operation is unsuitable using a computer with fixed color table 
H_ERR_NFSC = 5168
#: Computer represents gray scales only 
H_ERR_NACD = 5169
#: LUT of this display is full 
H_ERR_LUTO = 5170
#: Internal error: wrong color code 
H_ERR_WCC = 5171
#: Wrong type for window attribute 
H_ERR_WWATTRT = 5172
#: Wrong name for window attribute 
H_ERR_WWATTRN = 5173
#: negative height of area (or 0) 
H_ERR_WRSPART = 5174
#: negative width of area (or 0) 
H_ERR_WCSPART = 5175
#: Window not completely visible 
H_ERR_WNCV = 5176
#: Font not allowed for this operation 
H_ERR_FONT_NA = 5177
#: Window was created in different thread 
H_ERR_WDIFFTH = 5178
#: Drawing object already attached to another window 
H_ERR_OBJ_ATTACHED = 5194
#: Internal error: only RGB-Mode 
H_ERR_CHA3 = 5180
#: No more (image-)windows available 
H_ERR_NMWA = 5181
#: Depth was not stored with window 
H_ERR_DEPTH_NOT_STORED = 5179
#: Object index was not stored with window 
H_ERR_INDEX_NOT_STORED = 5182
#: Operator does not support primitives without point coordinates 
H_ERR_PRIM_NO_POINTS = 5183
#: Maximum image size for Windows Remote Desktop exceeded 
H_ERR_REMOTE_DESKTOP_SIZE = 5184
#: No OpenGL support available 
H_ERR_NOGL = 5185
#: No depth information available 
H_ERR_NODEPTH = 5186
#: OpenGL error 
H_ERR_OGL_ERROR = 5187
#: Required framebuffer object is unsupported 
H_ERR_UNSUPPORTED_FBO = 5188
#: OpenGL accelerated hidden surface removal not supported on this machine 
H_ERR_OGL_HSR_NOT_SUPPORTED = 5189
#: Invalid window parameter 
H_ERR_WP_IWP = 5190
#: Invalid value for window parameter 
H_ERR_WP_IWPV = 5191
#: Unknown mode 
H_ERR_UMOD = 5192
#: No image attached 
H_ERR_ATTIMG = 5193
#: invalid navigation mode 
H_ERR_NVG_WM = 5195
#: Internal file error 
H_ERR_FINTERN = 5196
#: Error while file synchronisation 
H_ERR_FS = 5197
#: insufficient rights 
H_ERR_FISR = 5198
#: Bad file descriptor 
H_ERR_BFD = 5199
#: File not found 
H_ERR_FNF = 5200
#: Error while writing image data (sufficient memory ?) 
H_ERR_DWI = 5201
#: Error while writing image descriptor (sufficient memory ?) 
H_ERR_DWID = 5202
#: Error while reading image data (format of image too small ?) 
H_ERR_DRI1 = 5203
#: Error while reading image data (format of image too big ?) 
H_ERR_DRI2 = 5204
#: Error while reading image descriptor: File too small 
H_ERR_DRID1 = 5205
#: Image matrices are different 
H_ERR_DIMMAT = 5206
#: Help file not found (setenv HALCONROOT) 
H_ERR_HNF = 5207
#: Help index not found (setenv HALCONROOT) 
H_ERR_XNF = 5208
#: File <standard_input> can not be closed 
H_ERR_CNCSI = 5209
#: <standard_output/error> can not be closed 
H_ERR_CNCSO = 5210
#: File can not be closed 
H_ERR_CNCF = 5211
#: Error while writing to file 
H_ERR_EDWF = 5212
#: Exceeding of maximum number of files 
H_ERR_NFA = 5213
#: Wrong file name 
H_ERR_WFIN = 5214
#: Error while opening the file 
H_ERR_CNOF = 5215
#: Wrong file mode 
H_ERR_WFMO = 5216
#: Wrong type for pixel (e.g. byte) 
H_ERR_WPTY = 5217
#: Wrong image width (too big ?) 
H_ERR_WIW = 5218
#: Wrong image height (too big ?) 
H_ERR_WIH = 5219
#: File already exhausted before reading an image 
H_ERR_FTS1 = 5220
#: File exhausted before terminating the image 
H_ERR_FTS2 = 5221
#: Wrong value for resolution (dpi) 
H_ERR_WDPI = 5222
#: Wrong output image size (width) 
H_ERR_WNOW = 5223
#: Wrong output image size (height) 
H_ERR_WNOH = 5224
#: Wrong number of parameter values: Format description 
H_ERR_WNFP = 5225
#: Wrong parameter name for operator 
H_ERR_WPNA = 5226
#: Wrong slot name for parameter 
H_ERR_WSNA = 5227
#: Operator class is missing in help file 
H_ERR_NPCF = 5228
#: Wrong or inconsistent help/ *.idx or help/ *.sta 
H_ERR_WHIF = 5229
#: File help/ *.idx not found 
H_ERR_HINF = 5230
#: File help/ *.sta not found 
H_ERR_HSNF = 5231
#: Inconsistent file help/ *.sta 
H_ERR_ICSF = 5232
#: No explication file (.exp) found 
H_ERR_EFNF = 5233
#: No file found in known graphic format 
H_ERR_NFWKEF = 5234
#: Wrong graphic format 
H_ERR_WIFT = 5235
#: Inconsistent file halcon.num 
H_ERR_ICNF = 5236
#: File with extension 'tiff' is no Tiff-file 
H_ERR_WTIFF = 5237
#: Wrong file format 
H_ERR_WFF = 5238
#: No PNM format 
H_ERR_NOPNM = 5242
#: Inconsistent or old help file 
H_ERR_ICODB = 5243
#: Invalid file encoding 
H_ERR_INVAL_FILE_ENC = 5244
#: File not open 
H_ERR_FNO = 5245
#: No files in use so far (none opened) 
H_ERR_NO_FILES = 5246
#: Invalid file format for regions 
H_ERR_NORFILE = 5247
#: Error while reading region data: Format of region too big. 
H_ERR_RDTB = 5248
#: Encoding for binary files not allowed 
H_ERR_BINFILE_ENC = 5249
#: Error reading file 
H_ERR_EDRF = 5250
#: Serial port not open 
H_ERR_SNO = 5251
#: No serial port available 
H_ERR_NSA = 5252
#: Could not open serial port 
H_ERR_CNOS = 5253
#: Could not close serial port 
H_ERR_CNCS = 5254
#: Could not get serial port attributes 
H_ERR_CNGSA = 5255
#: Could not set serial port attributes 
H_ERR_CNSSA = 5256
#: Wrong baud rate for serial connection 
H_ERR_WRSBR = 5257
#: Wrong number of data bits for serial connection 
H_ERR_WRSDB = 5258
#: Wrong flow control for serial connection 
H_ERR_WRSFC = 5259
#: Could not flush serial port 
H_ERR_CNFS = 5260
#: Error during write to serial port 
H_ERR_EDWS = 5261
#: Error during read from serial port 
H_ERR_EDRS = 5262
#: Serialized item does not contain valid regions. 
H_ERR_REG_NOSITEM = 5270
#: The version of the regions is not supported. 
H_ERR_REG_WRVERS = 5271
#: Serialized item does not contain valid images. 
H_ERR_IMG_NOSITEM = 5272
#: The version of the images is not supported. 
H_ERR_IMG_WRVERS = 5273
#: Serialized item does not contain valid XLD objects. 
H_ERR_XLD_NOSITEM = 5274
#: The version of the XLD objects is not supported. 
H_ERR_XLD_WRVERS = 5275
#: Serialized item does not contain valid objects. 
H_ERR_OBJ_NOSITEM = 5276
#: The version of the objects is not supported. 
H_ERR_OBJ_WRVERS = 5277
#: XLD object data can only be read by HALCON XL 
H_ERR_XLD_DATA_TOO_LARGE = 5678
#: Unexpected object detected 
H_ERR_OBJ_UNEXPECTED = 5279
#: File has not been opened in text format 
H_ERR_FNOTF = 5280
#: File has not been opened in binary file format 
H_ERR_FNOBF = 5281
#: Cannot create directory 
H_ERR_DIRCR = 5282
#: Cannot remove directory 
H_ERR_DIRRM = 5283
#: Cannot get current directory 
H_ERR_GETCWD = 5284
#: Cannot set current directory 
H_ERR_SETCWD = 5285
#: Need to call XInitThreads() 
H_ERR_XINIT = 5286
#: No image acquisition device opened 
H_ERR_NFS = 5300
#: IA: wrong color depth 
H_ERR_FGWC = 5301
#: IA: wrong device 
H_ERR_FGWD = 5302
#: IA: determination of video format not possible 
H_ERR_FGVF = 5303
#: IA: no video signal 
H_ERR_FGNV = 5304
#: Unknown image acquisition device 
H_ERR_UFG = 5305
#: IA: failed grabbing of an image 
H_ERR_FGF = 5306
#: IA: wrong resolution chosen 
H_ERR_FGWR = 5307
#: IA: wrong image part chosen 
H_ERR_FGWP = 5308
#: IA: wrong pixel ratio chosen 
H_ERR_FGWPR = 5309
#: IA: handle not valid 
H_ERR_FGWH = 5310
#: IA: instance not valid (already closed?) 
H_ERR_FGCL = 5311
#: Image acquisition device could not be initialized 
H_ERR_FGNI = 5312
#: IA: external triggering not supported 
H_ERR_FGET = 5313
#: IA: wrong camera input line (multiplex) 
H_ERR_FGLI = 5314
#: IA: wrong color space 
H_ERR_FGCS = 5315
#: IA: wrong port 
H_ERR_FGPT = 5316
#: IA: wrong camera type 
H_ERR_FGCT = 5317
#: IA: maximum number of acquisition device classes exceeded 
H_ERR_FGTM = 5318
#: IA: device busy 
H_ERR_FGDV = 5319
#: IA: asynchronous grab not supported 
H_ERR_FGASYNC = 5320
#: IA: unsupported parameter 
H_ERR_FGPARAM = 5321
#: IA: timeout 
H_ERR_FGTIMEOUT = 5322
#: IA: invalid gain 
H_ERR_FGGAIN = 5323
#: IA: invalid field 
H_ERR_FGFIELD = 5324
#: IA: invalid parameter type 
H_ERR_FGPART = 5325
#: IA: invalid parameter value 
H_ERR_FGPARV = 5326
#: IA: function not supported 
H_ERR_FGFNS = 5327
#: IA: incompatible interface version 
H_ERR_FGIVERS = 5328
#: IA: could not set parameter value 
H_ERR_FGSETPAR = 5329
#: IA: could not query parameter setting 
H_ERR_FGGETPAR = 5330
#: IA: parameter not available in current configuration 
H_ERR_FGPARNA = 5331
#: IA: device could not be closed properly 
H_ERR_FGCLOSE = 5332
#: IA: camera configuration file could not be opened 
H_ERR_FGCAMFILE = 5333
#: IA: unsupported callback type 
H_ERR_FGCALLBACK = 5334
#: IA: device lost 
H_ERR_FGDEVLOST = 5335
#: IA: grab aborted 
H_ERR_FGABORTED = 5336
#: IO: timeout 
H_ERR_IOTIMEOUT = 5350
#: IO: incompatible interface version 
H_ERR_IOIVERS = 5351
#: IO: handle not valid 
H_ERR_IOWH = 5352
#: IO: device busy 
H_ERR_IODBUSY = 5353
#: IO: insufficient user rights 
H_ERR_IOIAR = 5354
#: IO: device or channel not found 
H_ERR_IONF = 5355
#: IO: invalid parameter type 
H_ERR_IOPART = 5356
#: IO: invalid parameter value 
H_ERR_IOPARV = 5357
#: IO: invalid parameter number 
H_ERR_IOPARNUM = 5358
#: IO: unsupported parameter 
H_ERR_IOPARAM = 5359
#: IO: parameter not available in curr config.
H_ERR_IOPARNA = 5360
#: IO: function not supported 
H_ERR_IOFNS = 5361
#: IO: maximum number of dio classes exceeded
H_ERR_IOME = 5362
#: IO: driver of io device not available 
H_ERR_IODNA = 5363
#: IO: operation aborted 
H_ERR_IOABORTED = 5364
#: IO: invalid data type 
H_ERR_IODATT = 5365
#: IO: device lost 
H_ERR_IODEVLOST = 5366
#: IO: could not set parameter value 
H_ERR_IOSETPAR = 5367
#: IO: could not query parameter setting 
H_ERR_IOGETPAR = 5368
#: IO: device could not be closed properly 
H_ERR_IOCLOSE = 5369
#: Image type is not supported 
H_ERR_JXR_UNSUPPORTED_FORMAT = 5400
#: Invalid pixel format passed to filter function 
H_ERR_JXR_INVALID_PIXEL_FORMAT = 5401
#: Internal JpegXR error. 
H_ERR_JXR_INTERNAL_ERROR = 5402
#: Syntax error in output format string 
H_ERR_JXR_FORMAT_SYNTAX_ERROR = 5403
#: Maximum number of channels exceeded 
H_ERR_JXR_TOO_MANY_CHANNELS = 5404
#: Unspecified error in JXR library 
H_ERR_JXR_EC_ERROR = 5405
#: Bad magic number in JXR library 
H_ERR_JXR_EC_BADMAGIC = 5406
#: Feature not implemented in JXR library 
H_ERR_JXR_EC_FEATURE_NOT_IMPLEMENTED = 5407
#: File read/write error in JXR library 
H_ERR_JXR_EC_IO = 5408
#: Bad file format in JXR library 
H_ERR_JXR_EC_BADFORMAT = 5409
#: Error while closing the image file 
H_ERR_LIB_FILE_CLOSE = 5500
#: Error while opening the image file 
H_ERR_LIB_FILE_OPEN = 5501
#: Premature end of the image file 
H_ERR_LIB_UNEXPECTED_EOF = 5502
#: Image dimensions too large for this file format 
H_ERR_IDTL = 5503
#: Image too large for this HALCON version 
H_ERR_ITLHV = 5504
#: Too many iconic objects for this file format 
H_ERR_TMIO = 5505
#: File format is unsupported 
H_ERR_FILE_FORMAT_UNSUPPORTED = 5506
#: All channels must have equal dimensions 
H_ERR_INCONSISTENT_DIMENSIONS = 5507
#: File is no PCX-File 
H_ERR_PCX_NO_PCX_FILE = 5510
#: Unknown encoding 
H_ERR_PCX_UNKNOWN_ENCODING = 5511
#: More than 4 image plains 
H_ERR_PCX_MORE_THAN_4_PLANES = 5512
#: Wrong magic in color table 
H_ERR_PCX_COLORMAP_SIGNATURE = 5513
#: Wrong number of bytes in span 
H_ERR_PCX_REPEAT_COUNT_SPANS = 5514
#: Wrong number of bits/pixels 
H_ERR_PCX_TOO_MUCH_BITS_PIXEL = 5515
#: Wrong number of plains 
H_ERR_PCX_PACKED_PIXELS = 5516
#: File is no GIF-File 
H_ERR_GIF_NO_GIF_PICTURE = 5520
#: GIF: Wrong version 
H_ERR_GIF_BAD_VERSION = 5521
#: GIF: Wrong descriptor 
H_ERR_GIF_SCREEN_DESCRIPTOR = 5522
#: GIF: Wrong color table 
H_ERR_GIF_COLORMAP = 5523
#: GIF: Premature end of file 
H_ERR_GIF_READ_ERROR_EOF = 5524
#: GIF: Wrong number of images 
H_ERR_GIF_NOT_ENOUGH_IMAGES = 5525
#: GIF: Wrong image extension 
H_ERR_GIF_ERROR_ON_EXTENSION = 5526
#: GIF: Wrong left top width 
H_ERR_GIF_LEFT_TOP_WIDTH = 5527
#: GIF: Cyclic index of table 
H_ERR_GIF_CIRCULAR_TABL_ENTRY = 5528
#: GIF: Wrong image data 
H_ERR_GIF_BAD_IMAGE_DATA = 5529
#: File is no Sun-Raster-File 
H_ERR_SUN_RASTERFILE_TYPE = 5530
#: Wrong header 
H_ERR_SUN_RASTERFILE_HEADER = 5531
#: Wrong image width 
H_ERR_SUN_COLS = 5532
#: Wrong image height 
H_ERR_SUN_ROWS = 5533
#: Wrong color map 
H_ERR_SUN_COLORMAP = 5534
#: Wrong image data 
H_ERR_SUN_RASTERFILE_IMAGE = 5535
#: Wrong type of pixel 
H_ERR_SUN_IMPOSSIBLE_DATA = 5536
#: Wrong type of pixel 
H_ERR_XWD_IMPOSSIBLE_DATA = 5540
#: Wrong visual class 
H_ERR_XWD_VISUAL_CLASS = 5541
#: Wrong X10 header 
H_ERR_XWD_X10_HEADER = 5542
#: Wrong X11 header 
H_ERR_XWD_X11_HEADER = 5543
#: Wrong X10 colormap 
H_ERR_XWD_X10_COLORMAP = 5544
#: Wrong X11 colormap 
H_ERR_XWD_X11_COLORMAP = 5545
#: Wrong pixmap 
H_ERR_XWD_X11_PIXMAP = 5546
#: Unknown version 
H_ERR_XWD_UNKNOWN_VERSION = 5547
#: Error while reading an image 
H_ERR_XWD_READING_IMAGE = 5548
#: Error while reading a file 
H_ERR_TIF_BAD_INPUTDATA = 5550
#: Wrong colormap 
H_ERR_TIF_COLORMAP = 5551
#: Too many colors 
H_ERR_TIF_TOO_MANY_COLORS = 5552
#: Wrong photometric interpretation
H_ERR_TIF_BAD_PHOTOMETRIC = 5553
#: Wrong photometric depth 
H_ERR_TIF_PHOTOMETRIC_DEPTH = 5554
#: Image is no binary file 
H_ERR_TIF_NO_REGION = 5555
#: Unsupported TIFF format 
H_ERR_TIF_UNSUPPORTED_FORMAT = 5556
#: Wrong file format specification 
H_ERR_TIF_BAD_SPECIFICATION = 5557
#: TIFF file is corrupt 
H_ERR_TIF_FILE_CORRUPT = 5558
#: Required TIFF tag is missing 
H_ERR_TIF_TAG_UNDEFINED = 5559
#: File is no BMP-File 
H_ERR_BMP_NO_BMP_PICTURE = 5560
#: Premature end of file 
H_ERR_BMP_READ_ERROR_EOF = 5561
#: Incomplete header 
H_ERR_BMP_INCOMPLETE_HEADER = 5562
#: Unknown bitmap format 
H_ERR_BMP_UNKNOWN_FORMAT = 5563
#: Unknown compression format 
H_ERR_BMP_UNKNOWN_COMPRESSION = 5564
#: Wrong color table 
H_ERR_BMP_COLORMAP = 5565
#: Write error on output 
H_ERR_BMP_WRITE_ERROR = 5566
#: File does not contain a binary image 
H_ERR_BMP_NO_REGION = 5567
#: Wrong number of components in image 
H_ERR_JPG_COMP_NUM = 5570
#: Unknown error from libjpeg 
H_ERR_JPGLIB_UNKNOWN = 5571
#: Not implemented feature in libjpeg 
H_ERR_JPGLIB_NOTIMPL = 5572
#: File access error in libjpeg 
H_ERR_JPGLIB_FILE = 5573
#: Tmp file access error in libjpeg 
H_ERR_JPGLIB_TMPFILE = 5574
#: Memory error in libjpeg 
H_ERR_JPGLIB_MEMORY = 5575
#: Error in input image 
H_ERR_JPGLIB_INFORMAT = 5576
#: File is not a PNG file 
H_ERR_PNG_NO_PNG_FILE = 5580
#: Unknown interlace type 
H_ERR_PNG_UNKNOWN_INTERLACE_TYPE = 5581
#: Unsupported color type 
H_ERR_PNG_UNSUPPORTED_COLOR_TYPE = 5582
#: Image is no binary file 
H_ERR_PNG_NO_REGION = 5583
#: Image size too big 
H_ERR_PNG_SIZE_TOO_BIG = 5584
#: Error accessing TIFF tag 
H_ERR_TIF_TAG_ACCESS = 5587
#: Invalid TIFF tag value datatype 
H_ERR_TIF_TAG_DATATYPE = 5588
#: Unsupported TIFF tag requested 
H_ERR_TIF_TAG_UNSUPPORTED = 5589
#: File corrupt 
H_ERR_JP2_CORRUPT = 5590
#: Image precision too high 
H_ERR_JP2_PREC_TOO_HIGH = 5591
#: Error while encoding 
H_ERR_JP2_ENCODING_ERROR = 5592
#: Image size too big 
H_ERR_JP2_SIZE_TOO_BIG = 5593
#: Unknown internal error from libjasper 
H_ERR_JP2_INTERNAL_ERROR = 5594
#: File does not contain only images 
H_ERR_HOBJ_NOT_ONLY_IMAGES = 5599
#: Socket can not be set to block 
H_ERR_SOCKET_BLOCK = 5600
#: Socket can not be set to unblock 
H_ERR_SOCKET_UNBLOCK = 5601
#: Received data is no tuple 
H_ERR_SOCKET_NO_CPAR = 5602
#: Received data is no image 
H_ERR_SOCKET_NO_IMAGE = 5603
#: Received data is no region 
H_ERR_SOCKET_NO_RL = 5604
#: Received data is no xld object 
H_ERR_SOCKET_NO_XLD = 5605
#: Error while reading from socket 
H_ERR_SOCKET_READ_DATA_FAILED = 5606
#: Error while writing to socket 
H_ERR_SOCKET_WRITE_DATA_FAILED = 5607
#: Illegal number of bytes with get_rl 
H_ERR_SOCKET_WRONG_BYTE_NUMBER = 5608
#: Buffer overflow in read_data 
H_ERR_SOCKET_BUFFER_OVERFLOW = 5609
#: Socket can not be created 
H_ERR_SOCKET_CANT_ASSIGN_FD = 5610
#: Bind on socket failed 
H_ERR_SOCKET_CANT_BIND = 5611
#: Socket information is not available 
H_ERR_SOCKET_CANT_GET_PORTNUMBER = 5612
#: Socket cannot listen for incoming connections 
H_ERR_SOCKET_CANT_LISTEN = 5613
#: Connection could not be accepted 
H_ERR_SOCKET_CANT_ACCEPT = 5614
#: Connection request failed 
H_ERR_SOCKET_CANT_CONNECT = 5615
#: Hostname could not be resolved 
H_ERR_SOCKET_GETHOSTBYNAME = 5616
#: Unknown tuple type on socket 
H_ERR_SOCKET_ILLEGAL_TUPLE_TYPE = 5618
#: Timeout occurred on socket 
H_ERR_SOCKET_TIMEOUT = 5619
#: No more sockets available 
H_ERR_SOCKET_NA = 5620
#: Socket is not initialized 
H_ERR_SOCKET_NI = 5621
#: Invalid socket 
H_ERR_SOCKET_OOR = 5622
#: Socket is NULL 
H_ERR_SOCKET_IS = 5623
#: Received data type is too large 
H_ERR_SOCKET_DATA_TOO_LARGE = 5624
#: Wrong socket type. 
H_ERR_SOCKET_WRONG_TYPE = 5625
#: Received data is not packed. 
H_ERR_SOCKET_NO_PACKED_DATA = 5626
#: Socket parameter operation failed. 
H_ERR_SOCKET_PARAM_FAILED = 5627
#: The data does not match the format specification. 
H_ERR_SOCKET_FORMAT_MISMATCH = 5628
#: Invalid format specification. 
H_ERR_SOCKET_INVALID_FORMAT = 5629
#: Received data is no serialized item 
H_ERR_SOCKET_NO_SERIALIZED_ITEM = 5630
#: Unable to create SSL context 
H_ERR_SOCKET_TLS_CONTEXT = 5631
#: Invalid TLS certificate or private key 
H_ERR_SOCKET_TLS_CERT_KEY = 5632
#: Invalid TLS private key 
H_ERR_SOCKET_TLS_HANDSHAKE = 5633
#: Too many contours/polygons for this file format 
H_ERR_ARCINFO_TOO_MANY_XLDS = 5700
#: The version of the quaternion is not supported 
H_ERR_QUAT_WRONG_VERSION = 5750
#: Serialized item does not contain a valid quaternion
H_ERR_QUAT_NOSITEM = 5751
#: The version of the homogeneous matrix is not supported 
H_ERR_HOM_MAT2D_WRONG_VERSION = 5752
#: Serialized item does not contain a valid homogeneous matrix 
H_ERR_HOM_MAT2D_NOSITEM = 5753
#: The version of the homogeneous 3D matrix is not supported 
H_ERR_HOM_MAT3D_WRONG_VERSION = 5754
#: Serialized item does not contain a valid homogeneous 3D matrix 
H_ERR_HOM_MAT3D_NOSITEM = 5755
#: The version of the tuple is not supported 
H_ERR_TUPLE_WRONG_VERSION = 5756
#: Serialized item does not contain a valid tuple 
H_ERR_TUPLE_NOSITEM = 5757
#: Number too big for a string to number conversion (overflow) 
H_ERR_TUPLE_DTLFTHV = 5758
#: The version of the camera parameters (pose) is not supported 
H_ERR_POSE_WRONG_VERSION = 5759
#: Serialized item does not contain valid camera parameters (pose) 
H_ERR_POSE_NOSITEM = 5760
#: The version of the internal camera parameters is not supported 
H_ERR_CAM_PAR_WRONG_VERSION = 5761
#: Serialized item does not contain valid internal camera parameters 
H_ERR_CAM_PAR_NOSITEM = 5762
#: The version of the dual quaternion is not supported 
H_ERR_DUAL_QUAT_WRONG_VERSION = 5763
#: Serialized item does not contain a valid dual quaternion
H_ERR_DUAL_QUAT_NOSITEM = 5764
#: Image source operation failed - unknown reason 
H_ERR_IMGSRC_FAIL = 5800
#: Image source operation failed - wrong internal assumptions 
H_ERR_IMGSRC_LOGIC = 5801
#: Image source functionality is not implemented 
H_ERR_IMGSRC_NOT_IMPLEMENTED = 5802
#: Image source plugin version incompatible 
H_ERR_IMGSRC_INCOMPATIBLE_VERSION = 5803
#: Unhandled exception was triggered by a GenTL producer 
H_ERR_IMGSRC_GENTL_ERROR = 5804
#: Unhandled exception was triggered by the GenICam GenAPI 
H_ERR_IMGSRC_GENAPI_ERROR = 5805
#: Image source resource could not be initialized 
H_ERR_IMGSRC_RES_INIT_FAIL = 5806
#: Image source resource not initialized 
H_ERR_IMGSRC_RES_NOT_INITIALIZED = 5807
#: Image source module request is ambiguous 
H_ERR_IMGSRC_MOD_REQUEST_AMBIGUOUS = 5808
#: Image source module not found 
H_ERR_IMGSRC_MOD_NOT_FOUND = 5809
#: Image source parameter not found 
H_ERR_IMGSRC_PARAM_NOT_FOUND = 5810
#: Image source parameter - wrong value provided 
H_ERR_IMGSRC_PARAM_WRONG_VALUE = 5811
#: Image source parameter - wrong type provided 
H_ERR_IMGSRC_PARAM_WRONG_VALUE_TYPE = 5812
#: Image source parameter - value not readable 
H_ERR_IMGSRC_PARAM_VAL_NOT_READABLE = 5813
#: Image source parameter - value not writable 
H_ERR_IMGSRC_PARAM_VAL_NOT_WRITABLE = 5814
#: Image source parameter - property not available 
H_ERR_IMGSRC_PARAM_PROP_NOT_AVAILABLE = 5815
#: Image source parameter - command timeout 
H_ERR_IMGSRC_COMMAND_TIMEOUT = 5816
#: Image source operation failed - wrong internal state 
H_ERR_IMGSRC_WRONG_STATE = 5817
#: No images received within the configured timeout 
H_ERR_IMGSRC_FETCH_TIMEOUT = 5818
#: Waiting for images aborted 
H_ERR_IMGSRC_FETCH_ABORT = 5819
#: Pixel data conversion failed 
H_ERR_IMGSRC_CONVERSION_FAILED = 5820
#: Access to undefined memory area 
H_ERR_NP = 6000
#: Not enough memory available 
H_ERR_MEM = 6001
#: Memory partition on heap has been overwritten 
H_ERR_ICM = 6002
#: HAlloc: 0 bytes requested 
H_ERR_WMS = 6003
#: Tmp-memory management: Call freeing memory although nothing had been allocated 
H_ERR_NOTMP = 6004
#: Tmp-memory management: Null pointer while freeing 
H_ERR_TMPNULL = 6005
#: Tmp-memory management: Could not find memory element 
H_ERR_CNFMEM = 6006
#: memory management: wrong memory type 
H_ERR_WMT = 6007
#: Not enough video memory available 
H_ERR_MEM_VID = 6021
#: No memory block allocated at last 
H_ERR_NRA = 6041
#: System parameter for memory-allocation inconsistent 
H_ERR_IAD = 6040
#: Invalid alignment 
H_ERR_INVALID_ALIGN = 6042
#: Function was given a NULL ptr as input 
H_ERR_NULL_PTR = 6043
#: Process creation failed 
H_ERR_CP_FAILED = 6500
#: Wrong index for output control par. 
H_ERR_WOCPI = 7000
#: Wrong number of values: Output control parameter 
H_ERR_WOCPVN = 7001
#: Wrong type: Output control parameter 
H_ERR_WOCPT = 7002
#: Wrong data type for object key (input objects) 
H_ERR_WKT = 7003
#: Range for integer had been passed 
H_ERR_IOOR = 7004
#: Inconsistent Halcon version 
H_ERR_IHV = 7005
#: Not enough memory for strings allocated 
H_ERR_NISS = 7006
#: Internal error: Proc is NULL 
H_ERR_PROC_NULL = 7007
#: Unknown symbolic object key (input obj.) 
H_ERR_UNKN = 7105
#: Wrong number of output object parameter 
H_ERR_WOON = 7200
#: Output type <string> expected 
H_ERR_OTSE = 7400
#: Output type <long> expected 
H_ERR_OTLE = 7401
#: Output type <float> expected 
H_ERR_OTFE = 7402
#: Object parameter is a zero pointer 
H_ERR_OPINP = 7403
#: Tuple had been deleted; values are not valid any more 
H_ERR_TWC = 7404
#: CNN: Internal data error 
H_ERR_CNN_DATA = 7701
#: CNN: Invalid memory type 
H_ERR_CNN_MEM = 7702
#: CNN: Invalid data serialization 
H_ERR_CNN_IO_INVALID = 7703
#: CNN: Implementation not available 
H_ERR_CNN_IMPL_NOT_AVAILABLE = 7704
#: CNN: Wrong number of input data 
H_ERR_CNN_NUM_INPUTS_INVALID = 7705
#: CNN: Invalid implementation type 
H_ERR_CNN_IMPL_INVALID = 7706
#: CNN: Training is not supported in the current environment. 
H_ERR_CNN_TRAINING_NOT_SUP = 7707
#: For this operation a GPU with certain minimal requirements is required. See installation guide for details. 
H_ERR_CNN_GPU_REQUIRED = 7708
#: For this operation the CUDA library needs to be available. (See installation guide for details.) 
H_ERR_CNN_CUDA_LIBS_MISSING = 7709
#: OCR File: Error while reading classifier 
H_ERR_OCR_CNN_RE = 7710
#: Wrong generic parameter name 
H_ERR_OCR_CNN_WGPN = 7711
#: One of the parameters returns several values and has to be used exclusively 
H_ERR_OCR_CNN_EXCLUSIV_PARAM = 7712
#: Wrong generic parameter name 
H_ERR_CNN_WGPN = 7713
#: Invalid labels. 
H_ERR_CNN_INVALID_LABELS = 7714
#: OCR File: Wrong file version
H_ERR_OCR_CNN_FILE_WRONG_VERSION = 7715
#: Invalid classes: At least one class apears twice 
H_ERR_CNN_MULTIPLE_CLASSES = 7716
#: For this operation the cuBLAS library needs to be available. (See installation guide for details.) 
H_ERR_CNN_CUBLAS_LIBS_MISSING = 7717
#: For this operation the CUDNN library needs to be available. (See installation guide for details.) 
H_ERR_CNN_CUDNN_LIBS_MISSING = 7718
#: File 'find_text_support.hotc' not found (please place this file in the ocr subdirectory of the root directory of your HALCON installation or in the current working directory) 
H_ERR_OCR_FNF_FIND_TEXT_SUPPORT = 7719
#: Training step failed. This might be caused by unsuitable training parameters 
H_ERR_CNN_TRAINING_FAILED = 7720
#: Weights in Graph have been overwritten previously and are lost. 
H_ERR_CNN_NO_PRETRAINED_WEIGHTS = 7721
#: New input size is too small to produce meaningful features 
H_ERR_CNN_INVALID_INPUT_SIZE = 7722
#: Result is not available. 
H_ERR_CNN_RESULT_NOT_AVAILABLE = 7723
#: New number of channels must be either 1 or 3. 
H_ERR_CNN_INVALID_INPUT_DEPTH = 7724
#: New input number of channels can't be set to 3 if network is specified for number of channels 1 
H_ERR_CNN_DEPTH_NOT_AVAILABLE = 7725
#: Device batch size larger than batch size. 
H_ERR_CNN_INVALID_BATCH_SIZE = 7726
#: Invalid specification of a parameter. 
H_ERR_CNN_INVALID_PARAM_SPEC = 7727
#: Memory size exceeds maximal allowed value. 
H_ERR_CNN_EXCEEDS_MAX_MEM = 7728
#: New batch size causes integer overflow 
H_ERR_CNN_BATCH_SIZE_OVERFLOW = 7729
#: Invalid input image size for detection model 
H_ERR_CNN_INVALID_IMAGE_SIZE = 7730
#: Invalid parameter value for current layer 
H_ERR_CNN_INVALID_LAYER_PARAM_VALUE = 7731
#: Invalid parameter num for current layer 
H_ERR_CNN_INVALID_LAYER_PARAM_NUM = 7732
#: Invalid parameter type for current layer 
H_ERR_CNN_INVALID_LAYER_PARAM_TYPE = 7733
#: CNN: Wrong number of output data 
H_ERR_CNN_NUM_OUTPUTS_INVALID = 7734
#: CNN: Invalid input shape 
H_ERR_CNN_INVALID_SHAPE = 7735
#: CNN: Invalid input data 
H_ERR_CNN_INVALID_INPUT_DATA = 7736
#: CNN: For variable input lengths the ctc loss layer only computes correct gradients if the used cuDNN version is >= 7.6.3. Please upgrade cuDNN or do not use variable input lengths. 
H_ERR_CNN_CUDNN_CTC_LOSS_BUGGY = 7737
#: CNN: Invalid padding 
H_ERR_CNN_INVALID_PADDING = 7738
#: CNN: Invalid layer type serialization 
H_ERR_CNN_IO_INVALID_LAYER_TYPE = 7740
#: CNN: Inference failed 
H_ERR_CNN_INFERENCE_FAILED = 7741
#: CNN: Runtime not supported on this machine 
H_ERR_CNN_RUNTIME_FAILED = 7742
#: Graph: Internal error 
H_ERR_GRAPH_INTERNAL = 7751
#: Graph: Invalid data serialization 
H_ERR_GRAPH_IO_INVALID = 7752
#: Graph: Invalid index 
H_ERR_GRAPH_INVALID_INDEX = 7753
#: HCNNGraph: Internal error 
H_ERR_CNNGRAPH_INTERNAL = 7760
#: HCNNGraph: Invalid data serialization 
H_ERR_CNNGRAPH_IO_INVALID = 7761
#: HCNNGraph: Invalid layer specification 
H_ERR_CNNGRAPH_LAYER_INVALID = 7762
#: HCNNGraph: Graph not properly initialized 
H_ERR_CNNGRAPH_NOINIT = 7763
#: CNN-Graph: Invalid memory type 
H_ERR_CNNGRAPH_INVALID_MEM = 7764
#: CNN-Graph: Invalid number of layers 
H_ERR_CNNGRAPH_INVALID_NUML = 7765
#: CNN-Graph: Invalid index 
H_ERR_CNNGRAPH_INVALID_IDX = 7766
#: CNN-Graph: Invalid specification status 
H_ERR_CNNGRAPH_SPEC_STATUS = 7767
#: CNN-Graph: Graph is not allowed to be changed after initialization 
H_ERR_CNNGRAPH_NOCHANGE = 7768
#: CNN-Graph: Missing preprocessing 
H_ERR_CNNGRAPH_PREPROC = 7769
#: CNN-Graph: Invalid vertex degree 
H_ERR_CNNGRAPH_DEGREE = 7770
#: CNN-Graph: Invalid output shape 
H_ERR_CNNGRAPH_OUTSHAPE = 7771
#: CNN-Graph: Invalid specification 
H_ERR_CNNGRAPH_SPEC = 7772
#: CNN-Graph: Invalid graph definition 
H_ERR_CNNGRAPH_DEF = 7773
#: CNN-Graph: Architecture not suitable for the adaption of the number of output classes 
H_ERR_CNNGRAPH_NO_CLASS_CHANGE = 7774
#: CNN-Graph: Architecture not suitable for the adaption of the image size" 
H_ERR_CNNGRAPH_NO_IMAGE_RESIZE = 7775
#: CNN-Graph: Aux index out of bounds. 
H_ERR_CNNGRAPH_AUX_INDEX_OOB = 7776
#: CNN-Graph: Invalid graph definition. Probably the auxiliary outputs of a layer have not been connected with corresponding aux selection layers (SelectAux) or at least one aux output has not been specified during model creation (create_dl_model call). 
H_ERR_CNNGRAPH_AUX_SPEC = 7777
#: CNN-Graph: Layer not supported for selected runtime 
H_ERR_CNNGRAPH_LAYER_UNSUPPORTED = 7778
#: DL: Internal error 
H_ERR_DL_INTERNAL = 7779
#: DL: Error reading file 
H_ERR_DL_FILE_READ = 7780
#: DL: Error writing file 
H_ERR_DL_FILE_WRITE = 7781
#: DL: Wrong file version 
H_ERR_DL_FILE_WRONG_VERSION = 7782
#: DL: Inputs missing in input dict 
H_ERR_DL_INPUTS_MISSING = 7783
#: DL: Inputs have incorrect batch size 
H_ERR_DL_INPUT_WRONG_BS = 7784
#: DL: Invalid layer name 
H_ERR_DL_INVALID_NAME = 7785
#: DL: Duplicate layer name 
H_ERR_DL_DUPLICATE_NAME = 7786
#: DL: Invalid output layer 
H_ERR_DL_INVALID_OUTPUT = 7787
#: DL: Parameter is not available 
H_ERR_DL_PARAM_NOT_AVAILABLE = 7788
#: DL: Tuple inputs have incorrect length 
H_ERR_DL_INPUT_WRONG_LENGTH = 7789
#: DL: Tuple inputs have incorrect type 
H_ERR_DL_INPUT_WRONG_TYPE = 7790
#: DL: Some inputs have incorrect values 
H_ERR_DL_INPUT_WRONG_VALUES = 7791
#: DL: Some class ids are not unique 
H_ERR_DL_CLASS_IDS_NOT_UNIQUE = 7792
#: DL: Some class ids are invalid 
H_ERR_DL_CLASS_IDS_INVALID = 7793
#: DL: Input data of class id conversion is invalid. 
H_ERR_DL_CLASS_IDS_INVALID_CONV = 7794
#: DL: Type already defined 
H_ERR_DL_TYPE_ALREADY_DEFINED = 7795
#: DL: Cannot identify inference inputs. 
H_ERR_DL_NO_INFERENCE_INPUTS = 7796
#: DL: Some class ids overlap with ignore class ids. 
H_ERR_DL_CLASS_IDS_INVALID_OVERLAP = 7797
#: DL: Wrong number of output layer 
H_ERR_DL_WRONG_OUTPUT_LAYER_NUM = 7798
#: DL: Batch size multiplier needs to be greater than 0 
H_ERR_DL_WRONG_BS_MULTIPLIER = 7799
#: DL: Inputs have incorrect batch size. The number of needed inputs is defined by batch_size * batch_size_multiplier 
H_ERR_DL_INPUT_WRONG_BS_WITH_MULTIPLIER = 7800
#: Error occurred while reading an ONNX model 
H_ERR_DL_READ_ONNX = 7801
#: DL: Model does not have class ids 
H_ERR_DL_CLASS_IDS_MISSING = 7802
#: Error occurred while writing an ONNX model 
H_ERR_DL_WRITE_ONNX = 7803
#: DL: Libprotobuf for ONNX could not be loaded 
H_ERR_DL_ONNX_LOADER = 7804
#: DL: Wrong scales during FPN creation 
H_ERR_DL_FPN_SCALES = 7810
#: DL: Backbone unusable for FPN creation 
H_ERR_DL_FPN_INVALID_BACKBONE = 7811
#: DL: Backbone feature maps not divisible by 2 
H_ERR_DL_FPN_INVALID_FEATURE_MAP_SIZE = 7812
#: Invalid FPN levels given 
H_ERR_DL_FPN_INVALID_LEVELS = 7813
#: DL: Internal error using anchors 
H_ERR_DL_ANCHOR = 7820
#: DL: Invalid detector parameter 
H_ERR_DL_DETECTOR_INVALID_PARAM = 7821
#: DL: Invalid detector parameter value 
H_ERR_DL_DETECTOR_INVALID_PARAM_VALUE = 7822
#: DL: Invalid docking layer 
H_ERR_DL_DETECTOR_INVALID_DOCKING_LAYER = 7823
#: DL: Invalid instance type 
H_ERR_DL_DETECTOR_INVALID_INSTANCE_TYPE = 7824
#: DL-Node: Missing generic parameter 'name'. Please specify a layer name. 
H_ERR_DL_NODE_MISSING_PARAM_NAME = 7830
#: DL-Node: No generic parameter 'name' allowed for this node. 
H_ERR_DL_NODE_GENPARAM_NAME_NOT_ALLOWED = 7831
#: DL-Node: Invalid layer specification. 
H_ERR_DL_NODE_INVALID_SPEC = 7832
#: DL-Node: There can only be one direct connection between two layers.
H_ERR_DL_NODE_DUPLICATE_EDGE = 7833
#: DL-Solver: Invalid type. 
H_ERR_DL_SOLVER_INVALID_TYPE = 7840
#: DL-Solver: Invalid update formula. 
H_ERR_DL_SOLVER_INVALID_UPDATE_FORMULA = 7841
#: DL: Heatmap is unsupported with the selected runtime. 
H_ERR_DL_HEATMAP_UNSUPPORTED_RUNTIME = 7850
#: DL: Unsupported heatmap model type. The heatmap is only applicable for model type 'classification'. 
H_ERR_DL_HEATMAP_UNSUPPORTED_MODEL_TYPE = 7851
#: DL: Unsupported heatmap method 
H_ERR_DL_HEATMAP_UNSUPPORTED_METHOD = 7852
#: DL: Wrong target class id for heatmap 
H_ERR_DL_HEATMAP_WRONG_TARGET_CLASS_ID = 7853
#: DL: GC Anomaly Detection network not available 
H_ERR_DL_GCAD_NETWORK_NOT_AVAILABLE = 7870
#: DL: Internal error occurred in anomaly model 
H_ERR_DL_ANOMALY_MODEL_INTERNAL = 7880
#: DL: Untrained anomaly model 
H_ERR_DL_ANOMALY_MODEL_UNTRAINED = 7881
#: DL: Anomaly model training failed 
H_ERR_DL_ANOMALY_MODEL_TRAINING_FAILED = 7882
#: DL: Unable to set parameter on a trained anomaly detection model 
H_ERR_DL_ANOMALY_MODEL_PARAM_TRAINED = 7883
#: DL: Input image size cannot be changed 
H_ERR_DL_ANOMALY_MODEL_RESIZE = 7884
#: DL: Input depth is not supported 
H_ERR_DL_ANOMALY_MODEL_DEPTH = 7885
#: DL: Input domain must not be empty 
H_ERR_DL_ANOMALY_MODEL_INPUT_DOMAIN = 7886
#: Deep OCR internal error 
H_ERR_DEEP_OCR_MODEL_INTERNAL = 7890
#: Each entry of the alphabet can only contain a string of length one. 
H_ERR_DEEP_OCR_MODEL_INVALID_ALPHABET = 7891
#: Out of bounds index into alphabet. 
H_ERR_DEEP_OCR_MODEL_INVALID_ALPHABET_IDX = 7892
#: The type of the given DL model is not allowed. 
H_ERR_DEEP_OCR_MODEL_INVALID_MODEL_TYPE = 7893
#: The model is not available. 
H_ERR_DEEP_OCR_MODEL_NOT_AVAILABLE = 7894
#: It is not possible to specify a mapping because there is no internal alphabet specified. 
H_ERR_DEEP_OCR_MODEL_INVALID_ALPHABET_MAPPING_NO_ALPHABET = 7895
#: Out of bounds index into alphabet given as mapping. 
H_ERR_DEEP_OCR_MODEL_INVALID_ALPHABET_MAPPING_IDX = 7896
#: The length of the mapping and the length of the internal alphabet need to be the same. 
H_ERR_DEEP_OCR_MODEL_INVALID_ALPHABET_MAPPING_LEN = 7897
#: The model file cannot be found. 
H_ERR_DEEP_OCR_MODEL_FILE_NOT_FOUND = 7898
#: Some character is not part of the internal alphabet. 
H_ERR_DEEP_OCR_MODEL_UNKNOWN_CHAR = 7899
#: The given word length is invalid. 
H_ERR_DEEP_OCR_MODEL_INVALID_WORD_LENGTH = 7900
#: The given alphabet is not a unique list of characters 
H_ERR_DEEP_OCR_MODEL_ALPHABET_NOT_UNIQUE = 7901
#: apply_dl_model: no default outputs allowed 
H_ERR_DL_MODEL_APPLY_NO_DEF_OUTPUTS = 7910
#: DL: Unsupported generic parameter 
H_ERR_DL_MODEL_UNSUPPORTED_GENPARAM = 7911
#: DL: Operator does not support model 
H_ERR_DL_MODEL_OPERATOR_UNSUPPORTED = 7912
#: DL: Requested runtime cannot be set 
H_ERR_DL_MODEL_RUNTIME = 7913
#: DL: Unsupported generic value(s) 
H_ERR_DL_MODEL_UNSUPPORTED_GENVALUE = 7914
#: DL: Invalid number of samples 
H_ERR_DL_MODEL_INVALID_NUM_SAMPLES = 7915
#: DL: Parameter unsupported for converted model 
H_ERR_DL_MODEL_CONVERTED_PARAM = 7916
#: DL: Unsupported operation on converted model 
H_ERR_DL_MODEL_CONVERTED_UNSUPPORTED = 7917
#: DL: The given dataset is incorrect 
H_ERR_DL_INVALID_DATASET = 7925
#: DL: Invalid sample index 
H_ERR_DL_INVALID_SAMPLE_INDEX = 7926
#: Deep Counting model is not prepared 
H_ERR_DEEP_COUNTING_NOT_PREPARED = 7940
#: The chosen backbone is not settable 
H_ERR_DEEP_COUNTING_UNSUPPORTED_BACKBONE = 7941
#: Usage of prepare for a Deep Counting model is unsupported 
H_ERR_DEEP_COUNTING_PREPARE_UNSUPPORTED = 7942
#: Deep Counting model does not contain a backbone 
H_ERR_DEEP_COUNTING_NO_BACKBONE = 7943
#: DL: Unsupported device precision 
H_ERR_DL_DEVICE_UNSUPPORTED_PRECISION = 7960
#: DL: Pruning data does not fit the given model 
H_ERR_DL_PRUNING_WRONG_DATA = 7980
#: DL: Model architecture does not support pruning 
H_ERR_DL_PRUNING_UNSUPPORTED_BY_CNN = 7981
#: DL: Invalid model type for out-of-distribution detection 
H_ERR_DL_OOD_UNSUPPORTED_MODEL_TYPE = 7985
#: DL: Insufficient diverse samples for fitting out-of-distribution detection 
H_ERR_DL_OOD_INSUFFICIENT_SAMPLE_DIVERSITY = 7986
#: DL: Internal error in the calculation of out-of-distribution detection. 
H_ERR_DL_OOD_INTERNAL_ERROR = 7987
#: DL: Operation invalidates out-of-distribution detection. 
H_ERR_DL_OOD_INVALID = 7988
#: DLModule is not loaded 
H_ERR_DL_MODULE_NOT_LOADED = 7990
#: Unknown operator name 
H_ERR_WPRN = 8000
#: register_comp_used is not activated 
H_ERR_RCNA = 8001
#: Unknown operator class 
H_ERR_WPC = 8002
#: convol/mask: Error while opening file 
H_ERR_ORMF = 8101
#: convol/mask: Premature end of file 
H_ERR_EOFRMF = 8102
#: convol/mask: Conversion error 
H_ERR_CVTRMF = 8103
#: convol/mask: Wrong row-/column number 
H_ERR_LCNRMF = 8104
#: convol/mask: Mask size overflow 
H_ERR_WCOVRMF = 8105
#: convol/mask: Too many elements entered 
H_ERR_NEOFRMF = 8106
#: convol: Wrong margin type 
H_ERR_WRRA = 8107
#: convol: No mask object has got empty region 
H_ERR_MCN0 = 8108
#: convol: Weight factor is 0 
H_ERR_WF0 = 8110
#: convol: Inconsistent number of weights 
H_ERR_NWC = 8111
#: rank: Wrong rank value 
H_ERR_WRRV = 8112
#: convol/rank: Error while handling margin 
H_ERR_ROVFL = 8113
#: Error while parsing filter mask file 
H_ERR_EWPMF = 8114
#: Wrong number of coefficients for convolution (sigma too big?) 
H_ERR_WNUMM = 8120
#: No valid ID for data set 
H_ERR_WBEDN = 8200
#: No data set active (set_bg_esti) 
H_ERR_NBEDA = 8201
#: ID already used for data set 
H_ERR_BEDNAU = 8202
#: No data set created (create_bg_esti) 
H_ERR_NBEDC = 8204
#: Not possible to pass an object list 
H_ERR_NTM = 8205
#: Image has other size than the background image in data set 
H_ERR_WISBE = 8206
#: Up-date-region is bigger than background image 
H_ERR_UDNSSBE = 8207
#: Number of statistic data sets is too small 
H_ERR_SNBETS = 8208
#: Wrong value for adapt mode 
H_ERR_WAMBE = 8209
#: Wrong value for frame mode 
H_ERR_WFMBE = 8210
#: Number of point corresponcences too small 
H_ERR_PE_NPCTS = 8250
#: Invalid method 
H_ERR_PE_INVMET = 8251
#: Maximum number of fonts exceeded 
H_ERR_OCR_MEM1 = 8300
#: Wrong ID (Number) for font 
H_ERR_OCR_WID = 8301
#: OCR internal error: wrong ID 
H_ERR_OCR1 = 8302
#: OCR not initialised: no font was read in 
H_ERR_OCR_NNI = 8303
#: No font activated 
H_ERR_OCR_NAI = 8304
#: OCR internal error: Wrong threshold in angle determination 
H_ERR_OCR_WTP = 8305
#: OCR internal error: Wrong attribute 
H_ERR_OCR_WF = 8306
#: The version of the OCR classifier is not supported 
H_ERR_OCR_READ = 8307
#: OCR File: Inconsistent number of nodes 
H_ERR_OCR_NODES = 8308
#: OCR File: File too short 
H_ERR_OCR_EOF = 8309
#: OCR: Internal error 1 
H_ERR_OCR_INC1 = 8310
#: OCR: Internal error 2 
H_ERR_OCR_INC2 = 8311
#: Wrong type of OCR tool (no 'box' or 'net') 
H_ERR_WOCRTYPE = 8312
#: The version of the OCR training characters is not supported 
H_ERR_OCR_TRF = 8313
#: Image too large for training file 
H_ERR_TRF_ITL = 8314
#: Region too large for training file 
H_ERR_TRF_RTL = 8315
#: Protected OCR training file 
H_ERR_TRF_PT = 8316
#: Protected OCR training file: wrong passw. 
H_ERR_TRF_WPW = 8317
#: Serialized item does not contain a valid OCR classifier 
H_ERR_OCR_NOSITEM = 8318
#: OCR training file concatenation failed: identical input and output files 
H_ERR_TRF_CON_EIO = 8319
#: Invalid file format for MLP classifier 
H_ERR_OCR_MLP_NOCLASSFILE = 8320
#: The version of the MLP classifier is not supported 
H_ERR_OCR_MLP_WRCLASSVERS = 8321
#: Serialized item does not contain a valid MLP classifier 
H_ERR_OCR_MLP_NOSITEM = 8322
#: Invalid file format for SVM classifier 
H_ERR_OCR_SVM_NOCLASSFILE = 8330
#: The version of the SVM classifier is not supported
H_ERR_OCR_SVM_WRCLASSVERS = 8331
#: Serialized item does not contain a valid SVM classifier 
H_ERR_OCR_SVM_NOSITEM = 8332
#: Invalid file format for k-NN classifier 
H_ERR_OCR_KNN_NOCLASSFILE = 8333
#: Serialized item does not contain a valid k-NN classifier 
H_ERR_OCR_KNN_NOSITEM = 8334
#: Invalid file format for CNN classifier 
H_ERR_OCR_CNN_NOCLASSFILE = 8335
#: The version of the CNN classifier is not supported 
H_ERR_OCR_CNN_WRCLASSVERS = 8336
#: Serialized item does not contain a valid CNN classifier 
H_ERR_OCR_CNN_NOSITEM = 8337
#: Result name is not available for this mode 
H_ERR_OCR_RESULT_NOT_AVAILABLE = 8338
#: OCV system not initialized 
H_ERR_OCV_NI = 8350
#: The version of the OCV tool is not supported 
H_ERR_WOCVTYPE = 8351
#: Wrong name for an OCV object 
H_ERR_OCV_WNAME = 8353
#: Training has already been applied 
H_ERR_OCV_II = 8354
#: No training has been applied 
H_ERR_OCV_NOTTR = 8355
#: Serialized item does not contain a valid OCV tool 
H_ERR_OCV_NOSITEM = 8356
#: Wrong number of function points 
H_ERR_WLENGTH = 8370
#: List of values is not a function 
H_ERR_NO_FUNCTION = 8371
#: Wrong ordering of values (not ascending)
H_ERR_NOT_ASCENDING = 8372
#: Illegal distance of function points 
H_ERR_ILLEGAL_DIST = 8373
#: Function is not monotonic. 
H_ERR_NOT_MONOTONIC = 8374
#: Wrong function type. 
H_ERR_WFUNCTION = 8375
#: Same x-value due to double to float conversion. 
H_ERR_SAME_XVAL_CONV = 8376
#: The input points could not be arranged in a regular grid 
H_ERR_GRID_CONNECT_POINTS = 8390
#: Error while creating the output map 
H_ERR_GRID_GEN_MAP = 8391
#: Auto rotation failed 
H_ERR_GRID_AUTO_ROT = 8392
#: No common camera parameters 
H_ERR_CAL_NO_COMM_PAR = 8393
#: Vy must be > 0 
H_ERR_CAL_NEGVY = 8394
#: Same finder pattern found multiple times 
H_ERR_CAL_IDENTICAL_FP = 8395
#: Function not available for line scan cameras with perspective lenses 
H_ERR_CAL_LSCPNA = 8396
#: Mark segmentation failed 
H_ERR_CAL_MARK_SEGM = 8397
#: Contour extraction failed 
H_ERR_CAL_CONT_EXT = 8398
#: No finder pattern found 
H_ERR_CAL_NO_FP = 8399
#: At least 3 calibration points have to be indicated 
H_ERR_CAL_LCALP = 8400
#: Inconsistent finder pattern positions 
H_ERR_CAL_INCONSISTENT_FP = 8401
#: No calibration table found 
H_ERR_CAL_NCPF = 8402
#: Error while reading calibration table description file 
H_ERR_CAL_RECPF = 8403
#: Minimum threshold while searching for ellipses 
H_ERR_CAL_LTMTH = 8404
#: Read error / format error in calibration table description file 
H_ERR_CAL_FRCP = 8405
#: Error in projection: s_x = 0 or s_y = 0 or z = 0 
H_ERR_CAL_PROJ = 8406
#: Error in inverse projection 
H_ERR_CAL_UNPRO = 8407
#: Not possible to open camera parameter file 
H_ERR_CAL_RICPF = 8408
#: Format error in file: No colon 
H_ERR_CAL_FICP1 = 8409
#: Format error in file: 2. colon is missing 
H_ERR_CAL_FICP2 = 8410
#: Format error in file: Semicolon is missing 
H_ERR_CAL_FICP3 = 8411
#: Not possible to open camera parameter (pose) file 
H_ERR_CAL_REPOS = 8412
#: Format error in camera parameter (pose) file 
H_ERR_CAL_FOPOS = 8413
#: Not possible to open calibration target description file 
H_ERR_CAL_OCPDF = 8414
#: Not possible to open postscript file of calibration target 
H_ERR_CAL_OCPPS = 8415
#: Error while norming the vector 
H_ERR_CAL_EVECN = 8416
#: Fitting of calibration target failed 
H_ERR_CAL_NPLAN = 8417
#: No next mark found 
H_ERR_CAL_NNMAR = 8418
#: Normal equation system is not solvable 
H_ERR_CAL_NNEQU = 8419
#: Average quadratic error is too big for 3D position of mark 
H_ERR_CAL_QETHM = 8420
#: Non elliptic contour 
H_ERR_CAL_NOELL = 8421
#: Wrong parameter value slvand() 
H_ERR_CAL_WPARV = 8422
#: Wrong function results slvand() 
H_ERR_CAL_WFRES = 8423
#: Distance of marks in calibration target description file is not possible 
H_ERR_CAL_ECPDI = 8424
#: Specified flag for degree of freedom not valid 
H_ERR_CAL_WEFLA = 8425
#: Minimum error did not fall below 
H_ERR_CAL_NOMER = 8426
#: Wrong type in Pose (rotation / translation) 
H_ERR_CAL_WPTYP = 8427
#: Image size does not match the measurement in camera parameters 
H_ERR_CAL_WIMSZ = 8428
#: Point could not be projected into linescan image 
H_ERR_CAL_NPILS = 8429
#: Diameter of calibration marks could not be determined 
H_ERR_CAL_DIACM = 8430
#: Orientation of calibration plate could not be determined 
H_ERR_CAL_ORICP = 8431
#: Calibration plate does not lie completely inside the image 
H_ERR_CAL_CPNII = 8432
#: Wrong number of calibration marks extracted 
H_ERR_CAL_WNCME = 8433
#: Unknown name of parameter group 
H_ERR_CAL_UNKPG = 8434
#: Focal length must be non-negative 
H_ERR_CAL_NEGFL = 8435
#: Function not available for cameras with telecentric lenses 
H_ERR_CAL_TELNA = 8436
#: Function not available for line scan cameras 
H_ERR_CAL_LSCNA = 8437
#: Ellipse is degenerated to a point 
H_ERR_CAL_ELLDP = 8438
#: No orientation mark found 
H_ERR_CAL_NOMF = 8439
#: Camera calibration did not converge 
H_ERR_CAL_NCONV = 8440
#: Function not available for cameras with hypercentric lenses 
H_ERR_CAL_HYPNA = 8441
#: Point cannot be distorted. 
H_ERR_CAL_DISTORT = 8442
#: Wrong edge filter. 
H_ERR_CAL_WREDGFILT = 8443
#: Pixel size must be non-negative or zero 
H_ERR_CAL_NEGPS = 8444
#: Tilt is in the wrong range 
H_ERR_CAL_NEGTS = 8445
#: Rot is in the wrong range 
H_ERR_CAL_NEGRS = 8446
#: Camera parameters are invalid 
H_ERR_CAL_INVCAMPAR = 8447
#: Focal length must be positive 
H_ERR_CAL_ILLFL = 8448
#: Magnification must be positive 
H_ERR_CAL_ILLMAG = 8449
#: Illegal image plane distance 
H_ERR_CAL_ILLIPD = 8450
#: model not optimized yet - no res's
H_ERR_CM_NOT_OPTIMIZED = 8451
#: auxiliary model results not available 
H_ERR_CM_NOT_POSTPROCC = 8452
#: setup not 'visibly' interconnected 
H_ERR_CM_NOT_INTERCONN = 8453
#: camera parameter mismatch 
H_ERR_CM_CAMPAR_MISMCH = 8454
#: camera type mismatch 
H_ERR_CM_CAMTYP_MISMCH = 8455
#: camera type not supported 
H_ERR_CM_CAMTYP_UNSUPD = 8456
#: invalid camera ID 
H_ERR_CM_INVALD_CAMIDX = 8457
#: invalid cal.obj. ID 
H_ERR_CM_INVALD_DESCID = 8458
#: invalid cal.obj. instance ID 
H_ERR_CM_INVALD_COBJID = 8459
#: undefined camera 
H_ERR_CM_UNDEFINED_CAM = 8460
#: repeated observ. index 
H_ERR_CM_REPEATD_INDEX = 8461
#: undefined calib. object description 
H_ERR_CM_UNDEFI_CADESC = 8462
#: Invalid file format for calibration data model 
H_ERR_CM_NO_DESCR_FILE = 8463
#: The version of the calibration data model is not supported 
H_ERR_CM_WR_DESCR_VERS = 8464
#: zero-motion in linear scan camera 
H_ERR_CM_ZERO_MOTION = 8465
#: multi-camera and -calibobj not supported for all camera types 
H_ERR_CM_MULTICAM_UNSP = 8466
#: incomplete data, required for legacy calibration 
H_ERR_CM_INCMPLTE_DATA = 8467
#: Invalid file format for camera setup model 
H_ERR_CSM_NO_DESCR_FIL = 8468
#: The version of the camera setup model is not supported 
H_ERR_CSM_WR_DESCR_VER = 8469
#: full HALCON-caltab descr'n required 
H_ERR_CM_CALTAB_NOT_AV = 8470
#: invalid observation ID 
H_ERR_CM_INVAL_OBSERID = 8471
#: Serialized item does not contain a valid camera setup model 
H_ERR_CSM_NOSITEM = 8472
#: Serialized item does not contain a valid calibration data model 
H_ERR_CM_NOSITEM = 8473
#: Invalid tool pose id 
H_ERR_CM_INV_TOOLPOSID = 8474
#: Undefined tool pose 
H_ERR_CM_UNDEFINED_TOO = 8475
#: Invalid calib data model type 
H_ERR_CM_INVLD_MODL_TY = 8476
#: The camera setup model contains an uninitialized camera 
H_ERR_CSM_UNINIT_CAM = 8477
#: The hand-eye algorithm failed to find a solution. 
H_ERR_CM_NO_VALID_SOL = 8478
#: invalid observation pose 
H_ERR_CM_INVAL_OBS_POSE = 8479
#: Not enough calibration object poses 
H_ERR_CM_TOO_FEW_POSES = 8480
#: undefined camera type 
H_ERR_CM_UNDEF_CAM_TYP = 8481
#: Num of image pairs does not correspond to num of disparity values 
H_ERR_SM_INVLD_IMG_PAIRS_DISP_VAL = 8482
#: Invalid min/max disparity values 
H_ERR_SM_INVLD_DISP_VAL = 8483
#: No camera pair set by set_stereo_model_image_pairs 
H_ERR_SM_NO_IM_PAIR = 8484
#: No reconstructed point is visible for coloring 
H_ERR_SM_NO_VIS_COLOR = 8485
#: No camera pair yields reconstructed points (please check parameters of disparity method or bounding box) 
H_ERR_SM_NO_RECONSTRUCT = 8486
#: Partitioning of bounding box is too fine (please adapt the parameter 'resolution' or the bounding box)
H_ERR_SM_INVLD_BB_PARTITION = 8487
#: Invalid disparity range for binocular_disparity_ms method 
H_ERR_SM_INVLD_DISP_RANGE = 8488
#: Invalid param for binoculuar method 
H_ERR_SM_INVLD_BIN_PAR = 8489
#: invalid stereo model type 
H_ERR_SM_INVLD_MODL_TY = 8490
#: stereo model is not in persistent mode 
H_ERR_SM_NOT_PERSISTEN = 8491
#: invalid bounding box 
H_ERR_SM_INVLD_BOU_BOX = 8492
#: stereo reconstruction: image sizes must correspond to camera setup 
H_ERR_SR_INVLD_IMG_SIZ = 8493
#: bounding box is behind basis line 
H_ERR_SR_BBOX_BHND_CAM = 8494
#: Ambiguous calibration: Please, recalibrate with improved input data!
H_ERR_CAL_AMBIGUOUS = 8495
#: Pose of calibration plate not determined 
H_ERR_CAL_PCPND = 8496
#: Calibration failed: Please check your input data and calibrate again! 
H_ERR_CAL_FAILED = 8497
#: No observation data supplied! 
H_ERR_CAL_MISSING_DATA = 8498
#: The calibration object has to be seen at least once by every camera, if less than four cameras are used. 
H_ERR_CAL_FEWER_FOUR = 8499
#: Invalid file format for template 
H_ERR_NOAP = 8500
#: The version of the template is not supported 
H_ERR_WPFV = 8501
#: Error during changing the file mode 
H_ERR_MATCH_MODE = 8502
#: Inconsistent match file: Coordinates out of range 
H_ERR_MATCH_OOR = 8503
#: The image(s) is not a pyramid 
H_ERR_NOTAP = 8505
#: Number of template points too small 
H_ERR_NGTPTS = 8506
#: Template data can only be read by HALCON XL 
H_ERR_PDTL = 8507
#: Serialized item does not contain a valid NCC model 
H_ERR_NCC_NOSITEM = 8508
#: Serialized item does not contain a valid template 
H_ERR_MATCH_NOSITEM = 8509
#: Number of shape model points too small 
H_ERR_NTPTS = 8510
#: Gray and color shape models mixed 
H_ERR_CGSMM = 8511
#: Shape model data can only be read by HALCON XL 
H_ERR_SMTL = 8512
#: Shape model was not created from XLDs 
H_ERR_SMNXLD = 8513
#: Serialized item does not contain a valid shape model 
H_ERR_SM_NOSITEM = 8514
#: Shape model contour too near to clutter region 
H_ERR_SM_CL_CONT = 8515
#: Shape model does not contain clutter parameters 
H_ERR_SM_NO_CLUT = 8516
#: Shape models are not of the same clutter type 
H_ERR_SM_SAME_CL = 8517
#: Shape model has an invalid clutter contrast 
H_ERR_SM_WRONG_CLCO = 8518
#: Clutter region contains negative coordinates 
H_ERR_SM_CL_NEG = 8519
#: Box finder: Unsupported generic parameter 
H_ERR_FIND_BOX_UNSUP_GENPARAM = 8520
#: Initial components have different region types 
H_ERR_COMP_DRT = 8530
#: Solution of ambiguous matches failed 
H_ERR_COMP_SAMF = 8531
#: Computation of the incomplete gamma function not converged 
H_ERR_IGF_NC = 8532
#: Too many nodes while computing the minimum spanning arborescence 
H_ERR_MSA_TMN = 8533
#: Component training data can only be read by HALCON XL 
H_ERR_CTTL = 8534
#: Component model data can only be read by HALCON XL 
H_ERR_CMTL = 8535
#: Serialized item does not contain a valid component model 
H_ERR_COMP_NOSITEM = 8536
#: Serialized item does not contain a valid component training result 
H_ERR_TRAIN_COMP_NOSITEM = 8537
#: Size of the training image and the variation model differ 
H_ERR_VARIATION_WS = 8540
#: Variation model has not been prepared for segmentation 
H_ERR_VARIATION_PREP = 8541
#: Invalid variation model training mode 
H_ERR_VARIATION_WRMD = 8542
#: Invalid file format for variation model 
H_ERR_VARIATION_NOVF = 8543
#: The version of the variation model is not supported 
H_ERR_VARIATION_WVFV = 8544
#: Training data has been cleared 
H_ERR_VARIATION_TRDC = 8545
#: Serialized item does not contain a valid variation model 
H_ERR_VARIATION_NOSITEM = 8546
#: No more measure objects available 
H_ERR_MEASURE_NA = 8550
#: Measure object is not initialized 
H_ERR_MEASURE_NI = 8551
#: Invalid measure object 
H_ERR_MEASURE_OOR = 8552
#: Measure object is NULL 
H_ERR_MEASURE_IS = 8553
#: Measure object has wrong image size 
H_ERR_MEASURE_WS = 8554
#: Invalid file format for measure object 
H_ERR_MEASURE_NO_MODEL_FILE = 8555
#: The version of the measure object is not supported 
H_ERR_MEASURE_WRONG_VERSION = 8556
#: Measure object data can only be read by HALCON XL 
H_ERR_MEASURE_TL = 8557
#: Serialized item does not contain a valid measure object 
H_ERR_MEASURE_NOSITEM = 8558
#: Metrology model is not initialized 
H_ERR_METROLOGY_MODEL_NI = 8570
#: Invalid metrology object 
H_ERR_METROLOGY_OBJECT_INVALID = 8572
#: Not enough valid measures for fitting the metrology object 
H_ERR_METROLOGY_FIT_NOT_ENOUGH_MEASURES = 8573
#: Invalid file format for metrology model 
H_ERR_METROLOGY_NO_MODEL_FILE = 8575
#: The version of the metrology model is not supported 
H_ERR_METROLOGY_WRONG_VERSION = 8576
#: Fuzzy function is not set 
H_ERR_METROLOGY_NO_FUZZY_FUNC = 8577
#: Serialized item does not contain a valid metrology model 
H_ERR_METROLOGY_NOSITEM = 8578
#: Camera parameters are not set 
H_ERR_METROLOGY_UNDEF_CAMPAR = 8579
#: Pose of the measurement plane is not set 
H_ERR_METROLOGY_UNDEF_POSE = 8580
#: Mode of metrology model cannot be set since an object has already been added 
H_ERR_METROLOGY_SET_MODE = 8581
#: If the pose of the metrology object has been set several times, the operator is not longer allowed 
H_ERR_METROLOGY_OP_NOT_ALLOWED = 8582
#: All objects of a metrology model must have the same world pose and camera parameters. 
H_ERR_METROLOGY_MULTI_POSE_CAM_PAR = 8583
#: Input type of metrology model does not correspond with the current input type 
H_ERR_METROLOGY_WRONG_INPUT_MODE = 8584
#: Dynamic library could not be opened 
H_ERR_DLOPEN = 8600
#: Dynamic library could not be closed 
H_ERR_DLCLOSE = 8601
#: Symbol not found in dynamic library 
H_ERR_DLLOOKUP = 8602
#: Interface library not * available 
H_ERR_COMPONENT_NOT_INSTALLED = 8603
#: Not enough information for rad. calib. 
H_ERR_EAD_CAL_NII = 8650
#: The version of the shape model result is not supported 
H_ERR_WGSMFV = 8670
#: Restrict scale parameter outside the trained range 
H_ERR_GSM_INVALID_RES_SCALE = 8671
#: Angle parameter outside the trained range 
H_ERR_GSM_INVALID_ANGLE = 8672
#: Shape model needs training 
H_ERR_GSM_NEEDS_TRAINING = 8673
#: contrast_high cannot be smaller than contrast_low 
H_ERR_GSM_CONTRAST_HYS = 8674
#: Neither contrast_low nor contrast_high can be smaller than min_contrast 
H_ERR_GSM_CONTRAST_MIN_CONTRAST = 8675
#: iso_scale_max cannot be smaller than iso_scale_min 
H_ERR_GSM_ISO_SCALE_PAIR = 8676
#: scale_row_max cannot be smaller than scale_row_min 
H_ERR_GSM_ANISO_SCALE_ROW = 8677
#: scale_column_max cannot be smaller than scale_column_min 
H_ERR_GSM_ANISO_SCALE_COLUMN = 8678
#: Isotropic scaling not set 
H_ERR_GSM_ISO_NOT_SET = 8679
#: Anisotropic scaling not set 
H_ERR_GSM_ANISO_NOT_SET = 8680
#: No edge direction available to change shape matching metric 
H_ERR_GSM_INVALID_METRIC_XLD = 8681
#: Shape models with the same identifier cannot be searched simultaneously 
H_ERR_GSM_SAME_IDENTIFIER = 8682
#: Set parameters inconsistent with est. 'per_level' values 
H_ERR_SM_INCONSISTENT_PER_LEVEL = 8683
#: Extended parameter estimation failed 
H_ERR_GSM_EXT_PAR_EST = 8684
#: Wrong number of modules 
H_ERR_BAR_WNOM = 8701
#: Wrong number of elements 
H_ERR_BAR_WNOE = 8702
#: Unknown character (for this code) 
H_ERR_BAR_UNCHAR = 8703
#: Wrong name for attribute in barcode descriptor 
H_ERR_BAR_WRONGDESCR = 8705
#: Wrong thickness of element 
H_ERR_BAR_EL_LENGTH = 8706
#: No region found 
H_ERR_BAR_NO_REG = 8707
#: Wrong type of bar code 
H_ERR_BAR_WRONGCODE = 8708
#: Internal error in bar code reader 
H_ERR_BAR_INTERNAL = 8709
#: Candidate does not contain a decoded scanline 
H_ERR_BAR_NO_DECODED_SCANLINE = 8710
#: Empty model list 
H_ERR_BC_EMPTY_MODEL_LIST = 8721
#: Training cannot be done for multiple bar code types 
H_ERR_BC_TRAIN_ONLY_SINGLE = 8722
#: Cannot get bar code type specific parameter with get_bar_code_param. Use get_bar_code_param_specific 
H_ERR_BC_GET_SPECIFIC = 8723
#: Cannot get this object for multiple bar code types. Try again with single bar code type 
H_ERR_BC_GET_OBJ_MULTI = 8724
#: Wrong binary (file) format 
H_ERR_BC_WR_FILE_FORMAT = 8725
#: Wrong version of binary file 
H_ERR_BC_WR_FILE_VERS = 8726
#: The model must be in persistency mode to deliver the required object/result 
H_ERR_BC_NOT_PERSISTANT = 8727
#: Incorrect index of scanline's gray values
H_ERR_BC_GRAY_OUT_OF_RANGE = 8728
#: Neither find_bar_code nor decode_bar_code_rectanlge2 has been called in 'persistent' mode on this model 
H_ERR_NO_PERSISTENT_OP_CALL = 8729
#: The super-resolution algorithm has been aborted 
H_ERR_BC_ZOOMED_ABORTED = 8730
#: SRB: Invalid input data. 
H_ERR_BC_ZOOMED_INVALID_INPUT = 8731
#: Invalid input detected for barcode normalized cross correlation 
H_ERR_BC_XCORR_INVALID_INPUT = 8740
#: Too many bad rows found during barcode normalized cross correlation 
H_ERR_BC_XCORR_TOO_MANY_BAD_ROWS = 8741
#: No correlation found during barcode normalized cross correlation 
H_ERR_BC_XCORR_NO_CORRELATION = 8742
#: Invalid GS1 syntax dictionary 
H_ERR_INVALID_SYNTAX_DICTIONARY = 8743
#: Specified code type is not supported 
H_ERR_BAR2D_UNKNOWN_TYPE = 8800
#: Wrong foreground specified 
H_ERR_BAR2D_WRONG_FOREGROUND = 8801
#: Wrong matrix size specified 
H_ERR_BAR2D_WRONG_SIZE = 8802
#: Wrong symbol shape specified 
H_ERR_BAR2D_WRONG_SHAPE = 8803
#: Wrong generic parameter name 
H_ERR_BAR2D_WRONG_PARAM_NAME = 8804
#: Wrong generic parameter value 
H_ERR_BAR2D_WRONG_PARAM_VAL = 8805
#: Wrong symbol printing mode 
H_ERR_BAR2D_WRONG_MODE = 8806
#: Symbol region too near to image border 
H_ERR_BAR2D_SYMBOL_ON_BORDER = 8807
#: No rectangular module boundings found 
H_ERR_BAR2D_MODULE_CONT_NUM = 8808
#: Couldn't identify symbol finder 
H_ERR_BAR2D_SYMBOL_FINDER = 8809
#: Symbol region with wrong dimension 
H_ERR_BAR2D_SYMBOL_DIMENSION = 8810
#: Classification failed 
H_ERR_BAR2D_CLASSIF_FAILED = 8811
#: Decoding failed 
H_ERR_BAR2D_DECODING_FAILED = 8812
#: Reader programming not supported 
H_ERR_BAR2D_DECODING_READER = 8813
#: General 2d data code error 
H_ERR_DC2D_GENERAL = 8820
#: Corrupt signature of 2d data code handle 
H_ERR_DC2D_BROKEN_SIGN = 8821
#: Invalid 2d data code handle 
H_ERR_DC2D_INVALID_HANDLE = 8822
#: List of 2d data code models is empty 
H_ERR_DC2D_EMPTY_MODEL_LIST = 8823
#: Access to uninitialized (or not persistent) internal data 
H_ERR_DC2D_NOT_INITIALIZED = 8824
#: Invalid 'Candidate' parameter 
H_ERR_DC2D_INVALID_CANDIDATE = 8825
#: It's not possible to return more than one parameter for several candidates 
H_ERR_DC2D_INDEX_PARNUM = 8826
#: One of the parameters returns several values and has to be used exclusively for a single candidate 
H_ERR_DC2D_EXCLUSIV_PARAM = 8827
#: Parameter for default settings must be the first in the parameter list 
H_ERR_DC2D_DEF_SET_NOT_FIRST = 8828
#: Unexpected 2d data code error 
H_ERR_DC2D_INTERNAL_UNEXPECTED = 8829
#: Invalid parameter value 
H_ERR_DC2D_WRONG_PARAM_VALUE = 8830
#: Unknown parameter name 
H_ERR_DC2D_WRONG_PARAM_NAME = 8831
#: Invalid 'polarity' 
H_ERR_DC2D_WRONG_POLARITY = 8832
#: Invalid 'symbol_shape' 
H_ERR_DC2D_WRONG_SYMBOL_SHAPE = 8833
#: Invalid symbol size 
H_ERR_DC2D_WRONG_SYMBOL_SIZE = 8834
#: Invalid module size 
H_ERR_DC2D_WRONG_MODULE_SIZE = 8835
#: Invalid 'module_shape' 
H_ERR_DC2D_WRONG_MODULE_SHAPE = 8836
#: Invalid 'orientation' 
H_ERR_DC2D_WRONG_ORIENTATION = 8837
#: Invalid 'contrast_min' 
H_ERR_DC2D_WRONG_CONTRAST = 8838
#: Invalid 'measure_thresh' 
H_ERR_DC2D_WRONG_MEAS_THRESH = 8839
#: Invalid 'alt_measure_red' 
H_ERR_DC2D_WRONG_ALT_MEAS_RED = 8840
#: Invalid 'slant_max' 
H_ERR_DC2D_WRONG_SLANT = 8841
#: Invalid 'L_dist_max' 
H_ERR_DC2D_WRONG_L_DIST = 8842
#: Invalid 'L_length_min' 
H_ERR_DC2D_WRONG_L_LENGTH = 8843
#: Invalid module gap 
H_ERR_DC2D_WRONG_GAP = 8844
#: Invalid 'default_parameters' 
H_ERR_DC2D_WRONG_DEF_SET = 8845
#: Invalid 'back_texture' 
H_ERR_DC2D_WRONG_TEXTURED = 8846
#: Invalid 'mirrored' 
H_ERR_DC2D_WRONG_MIRRORED = 8847
#: Invalid 'classificator' 
H_ERR_DC2D_WRONG_CLASSIFICATOR = 8848
#: Invalid 'persistence' 
H_ERR_DC2D_WRONG_PERSISTENCE = 8849
#: Invalid model type 
H_ERR_DC2D_WRONG_MODEL_TYPE = 8850
#: Invalid 'module_roi_part' 
H_ERR_DC2D_WRONG_MOD_ROI_PART = 8851
#: Invalid 'finder_pattern_tolerance' 
H_ERR_DC2D_WRONG_FP_TOLERANCE = 8852
#: Invalid 'mod_aspect_max' 
H_ERR_DC2D_WRONG_MOD_ASPECT = 8853
#: Invalid 'small_modules_robustness' 
H_ERR_DC2D_WRONG_SM_ROBUSTNESS = 8854
#: Invalid 'contrast_tolerance' 
H_ERR_DC2D_WRONG_CONTRAST_TOL = 8855
#: Invalid 'alternating_pattern_tolerance' 
H_ERR_DC2D_WRONG_AP_TOLERANCE = 8856
#: Invalid 'deformation_tolerance' 
H_ERR_DC2D_WRONG_DEFORM_TOL = 8857
#: Invalid header in 2d data code model file 
H_ERR_DC2D_READ_HEAD_FORMAT = 8860
#: Invalid code signature in 2d data code model file 
H_ERR_DC2D_READ_HEAD_SIGN = 8861
#: Corrupted line in 2d data code model file 
H_ERR_DC2D_READ_LINE_FORMAT = 8862
#: Invalid module aspect ratio 
H_ERR_DC2D_WRONG_MODULE_ASPECT = 8863
#: wrong number of layers 
H_ERR_DC2D_WRONG_LAYER_NUM = 8864
#: wrong data code model version 
H_ERR_DCD_READ_WRONG_VERSION = 8865
#: Serialized item does not contain a valid 2D data code model 
H_ERR_DC2D_NOSITEM = 8866
#: Wrong binary (file) format 
H_ERR_DC2D_WR_FILE_FORMAT = 8867
#: Invalid parameter value 
H_ERR_SM3D_WRONG_PARAM_NAME = 8900
#: Invalid 'num_levels' 
H_ERR_SM3D_WRONG_NUM_LEVELS = 8901
#: Invalid 'optimization' 
H_ERR_SM3D_WRONG_OPTIMIZATION = 8902
#: Invalid 'metric' 
H_ERR_SM3D_WRONG_METRIC = 8903
#: Invalid 'min_face_angle' 
H_ERR_SM3D_WRONG_MIN_FACE_ANGLE = 8904
#: Invalid 'min_size' 
H_ERR_SM3D_WRONG_MIN_SIZE = 8905
#: Invalid 'model_tolerance' 
H_ERR_SM3D_WRONG_MODEL_TOLERANCE = 8906
#: Invalid 'fast_pose_refinment'
H_ERR_SM3D_WRONG_FAST_POSE_REF = 8907
#: Invalid 'lowest_model_level'
H_ERR_SM3D_WRONG_LOWEST_MODEL_LEVEL = 8908
#: Invalid 'part_size'
H_ERR_SM3D_WRONG_PART_SIZE = 8909
#: The projected model is too large (increase the value for DistMin or the image size in CamParam) 
H_ERR_SM3D_PROJECTION_TOO_LARGE = 8910
#: Invalid 'opengl_accuracy'
H_ERR_SM3D_WRONG_OPENGL_ACCURACY = 8911
#: Invalid 'recompute_score'
H_ERR_SM3D_WRONG_RECOMPUTE_SCORE = 8913
#: Invalid 'longitude_min' 
H_ERR_SM3D_WRONG_LON_MIN = 8920
#: Invalid 'longitude_max' 
H_ERR_SM3D_WRONG_LON_MAX = 8921
#: Invalid 'latitude_min 
H_ERR_SM3D_WRONG_LAT_MIN = 8922
#: Invalid 'latitude_max' 
H_ERR_SM3D_WRONG_LAT_MAX = 8923
#: Invalid 'cam_roll_min' 
H_ERR_SM3D_WRONG_ROL_MIN = 8924
#: Invalid 'cam_roll_max' 
H_ERR_SM3D_WRONG_ROL_MAX = 8925
#: Invalid 'dist_min' 
H_ERR_SM3D_WRONG_DIST_MIN = 8926
#: Invalid 'dist_max' 
H_ERR_SM3D_WRONG_DIST_MAX = 8927
#: Invalid 'num_matches' 
H_ERR_SM3D_WRONG_NUM_MATCHES = 8928
#: Invalid 'max_overlap' 
H_ERR_SM3D_WRONG_MAX_OVERLAP = 8929
#: Invalid 'pose_refinement' 
H_ERR_SM3D_WRONG_POSE_REFINEMENT = 8930
#: Invalid 'cov_pose_mode' 
H_ERR_SM3D_WRONG_COV_POSE_MODE = 8931
#: In. 'outlier_suppression' 
H_ERR_SM3D_WRONG_OUTLIER_SUP = 8932
#: Invalid 'border_model' 
H_ERR_SM3D_WRONG_BORDER_MODEL = 8933
#: Pose is not well-defined 
H_ERR_SM3D_UNDEFINED_POSE = 8940
#: Invalid file format for 3D shape model 
H_ERR_SM3D_NO_SM3D_FILE = 8941
#: The version of the 3D shape model is not supported 
H_ERR_SM3D_WRONG_FILE_VERSION = 8942
#: 3D shape model can only be read by HALCON XL 
H_ERR_SM3D_MTL = 8943
#: 3D object model does not contain any faces 
H_ERR_SM3D_NO_OM3D_FACES = 8944
#: Serialized item does not contain a valid 3D shape model 
H_ERR_SM3D_NOSITEM = 8945
#: Invalid 'union_adjacent_contours' 
H_ERR_SM3D_WRONG_UNION_ADJACENT_CONTOURS = 8946
#: Pose estimation model contains insufficient information 
H_ERR_DM3D_NO3DPOSEEST = 8947
#: Invalid file format for descriptor model 
H_ERR_DESCR_NODESCRFILE = 8960
#: The version of the descriptor model is not supported 
H_ERR_DESCR_WRDESCRVERS = 8961
#: Invalid 'radius' 
H_ERR_DM_WRONG_NUM_CIRC_RADIUS = 8962
#: Invalid 'check_neighbor' 
H_ERR_DM_WRONG_NUM_CHECK_NEIGH = 8963
#: Invalid 'min_check_neighbor_diff' 
H_ERR_DM_WRONG_NUM_MIN_CHECK_NEIGH = 8964
#: Invalid 'min_score' 
H_ERR_DM_WRONG_NUM_MIN_SCORE = 8965
#: Invalid 'sigma_grad' 
H_ERR_DM_WRONG_NUM_SIGMAGRAD = 8966
#: Invalid 'sigma_smooth' 
H_ERR_DM_WRONG_NUM_SIGMAINT = 8967
#: Invalid 'alpha' 
H_ERR_DM_WRONG_NUM_ALPHA = 8968
#: Invalid 'threshold' 
H_ERR_DM_WRONG_NUM_THRESHOLD = 8969
#: Invalid 'depth' 
H_ERR_DM_WRONG_NUM_DEPTH = 8970
#: Invalid 'number_trees' 
H_ERR_DM_WRONG_NUM_TREES = 8971
#: Invalid 'min_score_descr' 
H_ERR_DM_WRONG_NUM_MIN_SCORE_DESCR = 8972
#: Invalid 'patch_size' 
H_ERR_DM_WRONG_NUM_PATCH_SIZE = 8973
#: Invalid 'tilt' 
H_ERR_DM_WRONG_TILT = 8974
#: Invalid 'guided_matching' 
H_ERR_DM_WRONG_PAR_GUIDE = 8975
#: Invalid 'subpix' 
H_ERR_DM_WRONG_PAR_SUBPIX = 8976
#: Too few feature points can be found 
H_ERR_DM_TOO_FEW_POINTS = 8977
#: Invalid 'min_rot' 
H_ERR_DM_WRONG_NUM_MINROT = 8978
#: Invalid 'max_rot' 
H_ERR_DM_WRONG_NUM_MAXROT = 8979
#: Invalid 'min_scale' 
H_ERR_DM_WRONG_NUM_MINSCALE = 8980
#: Invalid 'max_scale' 
H_ERR_DM_WRONG_NUM_MAXSCALE = 8981
#: Invalid 'mask_size_grd' 
H_ERR_DM_WRONG_NUM_MASKSIZEGRD = 8982
#: Invalid 'mask_size_smooth' 
H_ERR_DM_WRONG_NUM_MASKSIZESMOOTH = 8983
#: Model broken 
H_ERR_BROKEN_MODEL = 8984
#: Invalid 'descriptor_type' 
H_ERR_DM_WRONG_DESCR_TYPE = 8985
#: Invalid 'matcher' 
H_ERR_DM_WRONG_PAR_MATCHER = 8986
#: Too many point classes - cannot be written to file 
H_ERR_DM_TOO_MANY_CLASSES = 8987
#: Serialized item does not contain a valid descriptor model 
H_ERR_DESCR_NOSITEM = 8988
#: Function not implemented on this machine 
H_ERR_NOT_IMPL = 9000
#: Image to process has wrong gray value type 
H_ERR_WIT = 9001
#: Wrong image component 
H_ERR_WIC = 9002
#: Undefined gray values 
H_ERR_UNDI = 9003
#: Wrong image format for operation (too big or too small) 
H_ERR_WIS = 9004
#: Wrong number of image components for image output 
H_ERR_WCN = 9005
#: String is too long (max. 1024 characters) 
H_ERR_STRTL = 9006
#: Wrong pixel type for this operation 
H_ERR_WITFO = 9007
#: Operation not realized yet for this pixel type 
H_ERR_NIIT = 9008
#: Image is no color image with three channels 
H_ERR_NOCIMA = 9009
#: Image acquisition devices are not supported in the demo version 
H_ERR_DEMO_NOFG = 9010
#: Packages are not supported in the demo version 
H_ERR_DEMO_NOPA = 9011
#: Internal Error: Unknown value
H_ERR_IEUNKV = 9020
#: Wrong parameter for this operation 
H_ERR_WPFO = 9021
#: Image domain too small 
H_ERR_IDTS = 9022
#: Draw operator has been canceled 
H_ERR_CNCLDRW = 9023
#: Error during matching of regular * expression 
H_ERR_REGEX_MATCH = 9024
#: Operator is not available in the student version of HALCON 
H_ERR_STUD_OPNA = 9050
#: Packages are not available in the student version of HALCON 
H_ERR_STUD_PANA = 9051
#: The selected image acquisition device is not available in the student version of HALCON
H_ERR_STUD_FGNA = 9052
#: No data points available 
H_ERR_NDPA = 9053
#: Object type is not supported. 
H_ERR_WR_OBJ_TYPE = 9054
#: Operator is disabled. 
H_ERR_OP_DISABLED = 9055
#: Too many unknown variables in linear equation 
H_ERR_TMU = 9100
#: No (unique) solution for the linear equation 
H_ERR_NUS = 9101
#: Too little equations in linear equation 
H_ERR_NEE = 9102
#: Points do not define a line 
H_ERR_PDDL = 9150
#: Matrix is not invertible 
H_ERR_MNI = 9200
#: Singular value decomposition did not converge 
H_ERR_SVD_CNVRG = 9201
#: Matrix has too few rows for singular value partition 
H_ERR_SVD_FEWROW = 9202
#: Eigenvalue computation did not converge 
H_ERR_TQLI_CNVRG = 9203
#: Eigenvalue computation did not converge 
H_ERR_JACOBI_CNVRG = 9204
#: Matrix is singular 
H_ERR_MATRIX_SING = 9205
#: Function matching did not converge 
H_ERR_MATCH_CNVRG = 9206
#: Input matrix undefined 
H_ERR_MAT_UNDEF = 9207
#: Input matrix with wrong dimension 
H_ERR_MAT_WDIM = 9208
#: Input matrix is not quadratic 
H_ERR_MAT_NSQR = 9209
#: Matrix operation failed 
H_ERR_MAT_FAIL = 9210
#: Matrix is not positive definite 
H_ERR_MAT_NPD = 9211
#: Matrix element division by 0 
H_ERR_MAT_DBZ = 9212
#: Matrix is not an upper triangular matrix 
H_ERR_MAT_NUT = 9213
#: Matrix is not a lower triangular matrix 
H_ERR_MAT_NLT = 9214
#: Matrix element is negative 
H_ERR_MAT_NEG = 9215
#: Matrix file: Invalid character 
H_ERR_MAT_UNCHAR = 9216
#: Matrix file: matrix incomplete 
H_ERR_MAT_NOT_COMPLETE = 9217
#: Invalid file format for matrix 
H_ERR_MAT_READ = 9218
#: Resulting matrix has complex values 
H_ERR_MAT_COMPLEX = 9219
#: Wrong value in matrix of exponents 
H_ERR_WMATEXP = 9220
#: The version of the matrix is not supported 
H_ERR_MAT_WRONG_VERSION = 9221
#: Serialized item does not contain a valid matrix 
H_ERR_MAT_NOSITEM = 9222
#: Internal Error: Wrong Node 
H_ERR_WNODE = 9230
#: Inconsistent red black tree 
H_ERR_CMP_INCONSISTENT = 9231
#: Internal error 
H_ERR_LAPACK_PAR = 9250
#: Number of points too small 
H_ERR_STRI_NPNT = 9260
#: First 3 points are collinear 
H_ERR_STRI_COLL = 9261
#: Identical points in triangulation 
H_ERR_STRI_IDPNT = 9262
#: Array not allocated large enough 
H_ERR_STRI_NALLOC = 9263
#: Triangle is degenerate 
H_ERR_STRI_DEGEN = 9264
#: Inconsistent triangulation 
H_ERR_STRI_ITRI = 9265
#: Self-intersecting polygon 
H_ERR_STRI_SELFINT = 9266
#: Inconsistent polygon data 
H_ERR_STRI_INCONS = 9267
#: Ambiguous great circle arc intersection 
H_ERR_STRI_AMBINT = 9268
#: Ambiguous great circle arc 
H_ERR_STRI_AMBARC = 9269
#: Illegal parameter 
H_ERR_STRI_ILLPAR = 9270
#: Not enough points for planar triangular meshing 
H_ERR_TRI_NPNT = 9280
#: The first three points of the triangular meshing are collinear 
H_ERR_TRI_COLL = 9281
#: Planar triangular meshing contains identical input points 
H_ERR_TRI_IDPNT = 9282
#: Invalid points for planar triangular meshing 
H_ERR_TRI_IDPNTIN = 9283
#: Internal error: allocated array too small for planar triangular meshing 
H_ERR_TRI_NALLOC = 9284
#: Internal error: planar triangular meshing inconsistent 
H_ERR_TRI_ITRI = 9285
#: Node index outside triangulation range 
H_ERR_TRI_OUTR = 9286
#: Local inconsistencies for all points with valid neighbors (parameters only allow few valid neighborhoods or point cloud not subsampled) 
H_ERR_TRI_LOCINC = 9290
#: Eye point and reference point coincide 
H_ERR_WSPVP = 9300
#: Real part of the dual quaternion has length 0 
H_ERR_DQ_ZERO_NORM = 9310
#: Timeout occurred 
H_ERR_TIMEOUT = 9400
#: Invalid 'timeout' 
H_ERR_WRONG_TIMEOUT = 9401
#: Invalid 'part_size' 
H_ERR_DEFORM_WRONG_NUM_CLUSTER = 9450
#: Invalid 'min_size' 
H_ERR_DEFORM_WRONG_NUM_MIN_SIZE = 9451
#: Invalid number of least-squares iterations 
H_ERR_DEFORM_WRONG_NUM_LSQ = 9452
#: Invalid 'angle_step' 
H_ERR_DEFORM_WRONG_ANGLE_STEP = 9453
#: Invalid 'scale_r_step' 
H_ERR_DEFORM_WRONG_SCALE_R_STEP = 9454
#: Invalid 'scale_c_step' 
H_ERR_DEFORM_WRONG_SCALE_C_STEP = 9455
#: Invalid 'max_angle_distortion' 
H_ERR_DEFORM_WRONG_MAX_ANGLE = 9456
#: Invalid 'max_aniso_scale_distortion' 
H_ERR_DEFORM_WRONG_MAX_ANISO = 9457
#: Invalid 'min_size' 
H_ERR_DEFORM_WRONG_MIN_SIZE = 9458
#: Invalid 'cov_pose_mode' 
H_ERR_DEFORM_WRONG_COV_POSE_MODE = 9459
#: Model contains no calibration information 
H_ERR_DEFORM_NO_CALIBRATION_INFO = 9460
#: Generic parameter name does not exist 
H_ERR_DEFORM_WRONG_PARAM_NAME = 9461
#: camera has different resolution than image 
H_ERR_DEFORM_IMAGE_TO_CAMERA_DIFF = 9462
#: Invalid file format for deformable model 
H_ERR_DEFORM_NO_MODEL_IN_FILE = 9463
#: The version of the deformable model is not supported 
H_ERR_DEFORM_WRONG_VERSION = 9464
#: Invalid 'deformation_smoothness'
H_ERR_DEFORM_WRONG_SMOOTH_DEFORM = 9465
#: Invalid 'expand_border' 
H_ERR_DEFORM_WRONG_EXPAND_BORDER = 9466
#: Model origin outside of axis-aligned bounding rectangle of template region 
H_ERR_DEFORM_ORIGIN_OUTSIDE_TEMPLATE = 9467
#: Serialized item does not contain a valid deformable model 
H_ERR_DEFORM_NOSITEM = 9468
#: Estimation of viewpose failed 
H_ERR_VIEW_ESTIM_FAIL = 9499
#: Object model has no points 
H_ERR_SFM_NO_POINTS = 9500
#: Object model has no faces 
H_ERR_SFM_NO_FACES = 9501
#: Object model has no normals 
H_ERR_SFM_NO_NORMALS = 9502
#: 3D surface model not trained for calculating view-based score 
H_ERR_SFM_NO_VISIBILITY = 9503
#: 3D surface model not trained for edge-supported matching 
H_ERR_SFM_NO_3D_EDGES = 9504
#: Invalid file format for 3D surface model 
H_ERR_SFM_NO_SFM_FILE = 9506
#: The version of the 3D surface model is not supported 
H_ERR_SFM_WRONG_FILE_VERSION = 9507
#: Serialized item does not contain a valid 3D surface model 
H_ERR_SFM_NOSITEM = 9508
#: Poses generate too many symmetries 
H_ERR_SFM_TOO_MANY_SYMMS = 9509
#: Invalid 3D file 
H_ERR_OM3D_INVALID_FILE = 9510
#: Invalid 3D Object Model 
H_ERR_OM3D_INVALID_MODEL = 9511
#: Unknown 3D file type 
H_ERR_OM3D_UNKNOWN_FILE_TYPE = 9512
#: The version of the 3D object model is not supported 
H_ERR_OM3D_WRONG_FILE_VERSION = 9513
#: Required attribute is missing 
H_ERR_OM3D_MISSING_ATTRIB = 9514
#: Required attribute point_coord is missing 
H_ERR_OM3D_MISSING_ATTRIB_V_COORD = 9515
#: Required attribute point_normal is missing 
H_ERR_OM3D_MISSING_ATTRIB_V_NORMALS = 9516
#: Required attribute face_triangle is missing 
H_ERR_OM3D_MISSING_ATTRIB_F_TRIANGLES = 9517
#: Required attribute line_array is missing 
H_ERR_OM3D_MISSING_ATTRIB_F_LINES = 9518
#: Required attribute f_trineighb is missing 
H_ERR_OM3D_MISSING_ATTRIB_F_TRINEIGB = 9519
#: Required attribute face_polygon is missing 
H_ERR_OM3D_MISSING_ATTRIB_F_POLYGONS = 9520
#: Required attribute xyz_mapping is missing 
H_ERR_OM3D_MISSING_ATTRIB_V_2DMAP = 9521
#: Required attribute o_primitive is missing 
H_ERR_OM3D_MISSING_ATTRIB_O_PRIMITIVE = 9522
#: Required attribute shape_model is missing 
H_ERR_OM3D_MISSING_ATTRIB_SHAPE_MODEL = 9523
#: Required extended attribute missing in 3D object model 
H_ERR_OM3D_MISSING_ATTRIB_EXTENDED = 9524
#: Serialized item does not contain a valid 3D object model 
H_ERR_OM3D_NOSITEM = 9525
#: Primitive in 3D object model has no extended data 
H_ERR_OM3D_MISSING_O_PRIMITIVE_EXTENSION = 9526
#: Operation invalid, 3D object model already contains triangles 
H_ERR_OM3D_CONTAIN_ATTRIB_F_TRIANGLES = 9527
#: Operation invalid, 3D object model already contains lines 
H_ERR_OM3D_CONTAIN_ATTRIB_F_LINES = 9528
#: Operation invalid, 3D object model already contains faces or polygons 
H_ERR_OM3D_CONTAIN_ATTRIB_F_POLYGONS = 9529
#: In a global registration an input object has no neighbors 
H_ERR_OM3D_ISOLATED_OBJECT = 9530
#: All components of points must be set at once 
H_ERR_OM3D_SET_ALL_COORD = 9531
#: All components of normals must be set at once 
H_ERR_OM3D_SET_ALL_NORMALS = 9532
#: Number of values doesn't correspond to number of already existing points 
H_ERR_OM3D_NUM_NOT_FIT_COORD = 9533
#: Number of values doesn't correspond to number of already existing normals 
H_ERR_OM3D_NUM_NOT_FIT_NORMALS = 9534
#: Number of values doesn't correspond to already existing triangulation 
H_ERR_OM3D_NUM_NOT_FIT_TRIANGLES = 9535
#: Number of values doesn't correspond to length of already existing polygons 
H_ERR_OM3D_NUM_NOT_FIT_POLYGONS = 9536
#: Number of values doesn't correspond to length of already existing polylines 
H_ERR_OM3D_NUM_NOT_FIT_LINES = 9537
#: Number of values doesn't correspond already existing 2D mapping 
H_ERR_OM3D_NUM_NOT_FIT_2DMAP = 9538
#: Number of values doesn't correspond to already existing extended attribute 
H_ERR_OM3D_NUM_NOT_FIT_EXTENDED = 9539
#: Per-face intensity is used with point attribute 
H_ERR_OM3D_FACE_INTENSITY_WITH_POINTS = 9540
#: Attribute is not (yet) supported 
H_ERR_OM3D_ATTRIBUTE_NOT_SUPPORTED = 9541
#: No point within bounding box 
H_ERR_OM3D_NOT_IN_BB = 9542
#: distance_in_front is smaller than the resolution 
H_ERR_DIF_TOO_SMALL = 9543
#: The minimum thickness is smaller than the surface tolerance 
H_ERR_MINTH_TOO_SMALL = 9544
#: Input width or height does not match the number of points in 3D object model 
H_ERR_OM3D_WRONG_DIMENSION = 9545
#: Image width or height must be set 
H_ERR_OM3D_MISSING_DIMENSION = 9546
#: Triangles of the 3D object model are not suitable for this operator 
H_ERR_SF_OM3D_TRIANGLES_NOT_SUITABLE = 9550
#: Too few suitable 3D points in the 3D object model 
H_ERR_SF_OM3D_FEW_POINTS = 9551
#: Not a valid serialized item file 
H_ERR_NO_SERIALIZED_ITEM = 9580
#: Serialized item: premature end of file 
H_ERR_END_OF_FILE = 9581
#: Invalid 'image_resize_method' 
H_ERR_SID_WRONG_RESIZE_METHOD = 9600
#: Invalid 'image_resize_value' 
H_ERR_SID_WRONG_RESIZE_VALUE = 9601
#: Invalid 'rating_method' 
H_ERR_SID_WRONG_RATING_METHOD = 9602
#: At least one type of image information must be used 
H_ERR_SID_NO_IMAGE_INFO_TYPE = 9603
#: Sample identifier does not contain color information 
H_ERR_SID_MODEL_NO_COLOR = 9604
#: Sample identifier does not contain texture information 
H_ERR_SID_MODEL_NO_TEXTURE = 9605
#: Sample image does not contain enough information 
H_ERR_SID_NO_IMAGE_INFO = 9606
#: Sample identifier does not contain unprepared data (use add_sample_identifier_preparation_data) 
H_ERR_SID_NO_UNPREPARED_DATA = 9607
#: Sample identifier has not been prepared yet (use prepare_sample_identifier) 
H_ERR_SID_MODEL_NOT_PREPARED = 9608
#: Sample identifier does not contain untrained data (use add_sample_identifier_training_data) 
H_ERR_SID_NO_UNTRAINED_DATA = 9609
#: Sample identifier has not been trained yet (use train_sample_identifier) 
H_ERR_SID_MODEL_NOT_TRAINED = 9610
#: Sample identifier does not contain result data 
H_ERR_SID_NO_RESULT_DATA = 9611
#: Sample identifier must contain at least two training objects (use add_sample_identifier_training_data) 
H_ERR_SID_NUM_TRAIN_OBJ = 9612
#: More than one user thread still uses HALCON * resources during finalization 
H_ERR_FINI_USR_THREADS = 9700
#: Invalid file format for encrypted items 
H_ERR_NO_ENCRYPTED_ITEM = 9800
#: Wrong password 
H_ERR_WRONG_PASSWORD = 9801
#: Encryption failed 
H_ERR_ENCRYPT_FAILED = 9802
#: Decryption failed 
H_ERR_DECRYPT_FAILED = 9803
#: User defined error codes must be larger than this value 
H_ERR_START_EXT = 10000
#: No license found 
H_ERR_NO_LICENSE = 2003
#: No modules in license (no VENDOR_STRING) 
H_ERR_NO_MODULES = 2005
#: No license for this operator 
H_ERR_NO_LIC_OPER = 2006
#: !HERRORDEF_H 
H_ERR_LAST_LIC_ERROR = H_ERR_LIC_NEWVER
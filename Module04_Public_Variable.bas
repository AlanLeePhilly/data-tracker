Attribute VB_Name = "Module04_Public_Variable"
'global variable used to find row to update
Public updateRow As Long

Public phaseHead As String
Public courtHead As String
Public tempHead As String
Public i As Long


Public Const selectedColor = &H8000000A
Public Const unselectedColor = &H8000000F

Public saveCounter As Integer
Public saveAsCounter As Integer

'global variable used to send textboxes to validators
Public ctl As Control

Public Lookup As Scripting.Dictionary

Public Const NUM_RESTITUTION_FILED_BUCKETS = 5
Public Const NUM_RESTITUTION_PAID_BUCKETS = 10
Public Const NUM_COURT_COST_FILED_BUCKETS = 5
Public Const NUM_COURT_COST_PAID_BUCKETS = 10
Public Const NUM_COMM_SERVICE_FILED_BUCKETS = 10
Public Const NUM_COMM_SERVICE_EARNED_BUCKETS = 20

Public Const NUM_SUPERVISION_BUCKETS_STANDARD = 15
Public Const NUM_SUPERVISION_BUCKETS_JTC = 30
Public Const NUM_SUPERVISION_BUCKETS_ADULT = 5
Public Const NUM_SUPERVISION_BUCKETS_AGG = 30
Public Const NUM_SUPERVISION_BUCKETS_AGG_AGG = 30

Public Const NUM_CONDITION_BUCKETS_STANDARD = 15
Public Const NUM_CONDITION_BUCKETS_JTC = 15
Public Const NUM_CONDITION_BUCKETS_ADULT = 5
Public Const NUM_CONDITION_BUCKETS_AGG = 20
Public Const NUM_CONDITION_BUCKETS_AGG_AGG = 20













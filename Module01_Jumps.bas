Attribute VB_Name = "Module01_Jumps"
Sub Petition()
    Range(headerFind("PETITION") & 2).Select
End Sub
Sub IntakeConference()
    Range(headerFind("INTAKE CONFERENCE") & 2).Select
End Sub
Sub Detention()
    Range(headerFind("DETENTION") & 2).Select
End Sub
Sub New_Detention()
    Range(headerFind("DETENTION (VOP)") & 2).Select
End Sub
Sub Diversion()
    Range(headerFind("DIVERSION") & 2).Select
End Sub
Sub FOUR_G()
    Range(headerFind("4G") & 2).Select
End Sub
Sub FOUR_E()
    Range(headerFind("4E") & 2).Select
End Sub
Sub SIX_F()
    Range(headerFind("6F") & 2).Select
End Sub
Sub SIX_H()
    Range(headerFind("6H") & 2).Select
End Sub
Sub THREE_E()
    Range(headerFind("3E") & 2).Select
End Sub
Sub Crossover()
    Range(headerFind("Crossover") & 2).Select
End Sub
Sub WRAP()
    Range(headerFind("WRAP") & 2).Select
End Sub
Sub JTC()
    Range(headerFind("JTC") & 2).Select
End Sub
Sub Adult()
    Range(headerFind("ADULT") & 2).Select
End Sub
Sub Aggregate_Pretrial()
    Range(hFind("Pretrial", "LEGAL STATUS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Consent_Decree()
    Range(hFind("Consent Decree", "LEGAL STATUS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Interim_Probation()
    Range(hFind("Interim Probation", "LEGAL STATUS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Probation()
    Range(hFind("Probation", "LEGAL STATUS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Aftercare_Probation()
    Range(hFind("Aftercare Probation", "LEGAL STATUS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Petition_Outcomes()
    Range(headerFind("Petition Outcomes") & 2).Select
End Sub
Sub Aggregate_Court_Proceedings()
    Range(hFind("COURT PROCEEDINGS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Certifications()
    Range(hFind("Certification", "COURT PROCEEDINGS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Admissions()
    Range(hFind("Admissions", "COURT PROCEEDINGS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Adjudications()
    Range(hFind("Adjudications", "COURT PROCEEDINGS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Continuances()
    Range(hFind("Continuances", "COURT PROCEEDINGS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Placements()
    Range(hFind("PLACEMENTS", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Supervision_Programs()
    Range(hFind("Supervision Programs", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_IPS()
    Range(hFind("Did Youth Have IPS?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Pre_ERC()
    Range(hFind("Did Youth Have Pre-ERC?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_IHD()
    Range(hFind("Did Youth Have IHD?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_ISP()
    Range(hFind("Did Youth Have ISP?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_GPS()
    Range(hFind("Did Youth Have GPS?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Post_ERC()
    Range(hFind("Did Youth Have Post-ERC?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Reintegration()
    Range(hFind("Did Youth Have Reintegration?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_CUA()
    Range(hFind("Did Youth Have CUA?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_RTF()
    Range(hFind("Did Youth Have RTF?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Inpatient_DA()
    Range(hFind("Did Youth Have Inpatient D&A?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Other_Supervision()
    Range(hFind("Did Youth Have Other Supervision?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Conditions()
    Range(hFind("Conditions", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_IOP()
    Range(hFind("Was Youth Ordered IOP?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Mental_Health()
    Range(headerFind("Was Youth Ordered Mental Health?") & 2).Select
End Sub
Sub Aggregate_Anger_Mgt()
    Range(hFind("Was Youth Ordered Anger Mgt.?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Curfew()
    Range(hFind("Was Youth Ordered Curfew?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_GED()
    Range(hFind("Was Youth Ordered GED?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Alternative_School()
    Range(hFind("Was Youth Ordered Alternative School?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Family_Therapy()
    Range(hFind("Was Youth Ordered Family Therapy?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Grief_Counseling()
    Range(hFind("Was Youth Ordered Grief Counseling?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Sexual_Counseling()
    Range(hFind("Was Youth Ordered Sexual Counseling?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_BHE()
    Range(hFind("Was Youth Ordered BHE?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Essay()
    Range(hFind("Was Youth Ordered Essay?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Victim_Conference()
    Range(hFind("Was Youth Ordered Victim Conference?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Other_Condition()
    Range(hFind("Was Youth Ordered Other Condition?", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Restitution()
    Range(hFind("Restitution", "AGGREGATES") & 2).Select
End Sub
Sub Aggregate_Comm_Service()
    Range(hFind("Comm. Service", "AGGREGATES") & 2).Select
End Sub
Sub Rearrests()
    Range(headerFind("Rearrests") & 2).Select
End Sub
Sub FTA()
    Range(headerFind("FTA") & 2).Select
End Sub
Sub Expungements()
    Range(hFind("Expungements", "AGGREGATES") & 2).Select
End Sub
Sub Phase2()
    Range(headerFind("PHASE II") & 2).Select
End Sub
Sub ReturntoNavigation()
    Range("A2").Select
End Sub
Sub Respites()
    Range("DP2").Select
End Sub
Sub Restitution()
    Range("EM2").Select
End Sub
Sub ListingHistory()
    Range(headerFind("LISTINGS") & 2).Select
End Sub
Sub NextRestitutionEntry()
    ' Jumps to next restitution entry
    ActiveWindow.SmallScroll ToRight:=-17
    ActiveCell.Offset(1, -25).Range("A1").Select
End Sub
Sub JumptoBlank()
    Range("C" & Rows.count).End(xlUp).Offset(1).Select
End Sub


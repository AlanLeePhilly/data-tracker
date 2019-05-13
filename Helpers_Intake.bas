Attribute VB_Name = "Helpers_Intake"
Sub aggFlag(ByVal userRow As Long)
    Call flagNo(Range(hFind("Was Youth on Pretrial?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Was Youth on Consent Decree?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Was Youth on Interim Probation?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Was Youth on Probation?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Was Youth on Aftercare Probation?", "AGGREGATES") & userRow))
    
    Call flagNo(Range(hFind("Was Notice of Certification Given?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Was Notice of De-Certification Given?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Did Youth Enter an Admission?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Adjudicated Delinquent?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Did Youth Have a Continuance?", "AGGREGATES") & userRow))
    
    Call flagNo(Range(hFind("Was Youth Placed?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Did Youth Have Restitution?", "AGGREGATES") & userRow))
    
    Call flagNo(Range(hFind("Did Youth Have Court Costs?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Did Youth Have Comm. Service?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Was Youth Rearrested?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Did Youth FTA?", "AGGREGATES") & userRow))
    Call flagNo(Range(hFind("Record Expunged?", "AGGREGATES") & userRow))
End Sub

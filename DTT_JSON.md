DTT data formats (all-caps fields are placeholders that have been changed to protect privacy)

From api/Disaster/&lt;DR ID&gt;/People/Details:  JSON array

```json
 {
    "Chapter": null,
    "DisasterGroupID": null,
    "Division": "D24",
    "Email": "EMAIL ADDRESS",
    "FirstName": "FIRST_NAME",
    "GAP": null,
    "Gender": "Female",
    "LastName": "LAST_NAME",
    "MobilePhone": "NNN-NNN-NNNN",
    "Name": "LAST_NAME, FIRST_NAME",
    "PersonID": 4315,
    "PersonType": null,
    "Region": null,
    "Responder": false,
    "Status": "Assign",
    "Transportation": "Rental",
    "UpdatedBy": "System",
    "UpdatedTimestamp": "2021-07-05T16:55:30.1210894-04:00",
    "VC": "VC_NUMBER",
    "WorkLocation": "Marriott Residence Inn Surfside",
    "tmStatus": null
  },
```

From api/Disaster/&lt;DR ID&gt;/Vehicles: JSON array

```json
 {
    "DisasterCode": "098-2021",
    "DisasterVehicleID": 5897,
    "Status": "Active",
    "TransferredToDRCode": null,
    "Vehicle": {
      "AssignedTo": null,
      "Attachments": null,
      "BoxTruckUsage": null,
      "Color": "Red",
      "CurrentDriverEffectiveDate": "2021-06-26T00:00:00-04:00",
      "CurrentDriverLodging": "Marriott - Biscyane Bay",
      "CurrentDriverName": "LAST_NAME, FIRST_NAME",
      "CurrentDriverNameAndMemberNo": "LAST_NAME, FIRST_NAME (VC_ID)",
      "CurrentDriverPersonId": 10275,
      "CurrentDriverWorkLocationId": null,
      "CurrentDriverWorkLocationName": "Marriott Residence Inn Surfside",
      "Deleted": false,
      "District": null,
      "Drivers": [],
      "DropoffAgencyId": null,
      "DueDate": "2021-07-25T00:00:00-04:00",
      "Exchanged": false,
      "Expiry": false,
      "ExtendedOn": null,
      "GAP": "MC/FF/MN",
      "KeyNumber": "96710213",
      "Make": "Toyota",
      "Model": " Corolla",
      "MotorPoolDistrict": null,
      "NewArrivalInspectionFiles": null,
      "NewOtherFiles": null,
      "NewPostInspectionFiles": null,
      "NewTowingFiles": null,
      "Notes": null,
      "OutProcessed": false,
      "PickupAgencyId": 868,
      "PickupAgencyName": "Avis Tampa Intl Airport",
      "Plate": "IEMX57",
      "PlateState": "FL",
      "PlateStateText": null,
      "RentalAgreementDropoffDate": null,
      "RentalAgreementNumber": "707143500",
      "RentalAgreementPerson": "LAST_NAME, FIRST_NAME",
      "RentalAgreementPersonId": 10275,
      "RentalAgreementPickupDate": "2021-06-25T00:00:00-04:00",
      "RentalAgreementReservationNumber": "06778641US2",
      "RentalAgreements": [],
      "Transferred": false,
      "VehicleAttachmentIdsToDelete": null,
      "VehicleCategoryCode": "R",
      "VehicleID": 5785,
      "VehicleType": "Car",
      "VehicleTypeID": 34,
      "Vendor": "Avis"
    }
  },
```

From api/Disaster/&lt;DR ID&gt;/agencies: JSON array

```json
[
  {
    "AgencyID": 336,
    "Name": "AVIS Sacramento CA APO",
    "Address": "6520 McNair Circle",
    "City": "Sacramento",
    "State": "CA",
    "StateText": null,
    "Zip": "95837",
    "Telephone": "916-922-5601",
    "DisasterID": 0,
    "UpdatedTimestamp": "2020-10-26T23:18:16.2179331-04:00",
    "UpdatedBy": "EMAIL_ADDRESS"
  },
]
```

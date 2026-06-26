<#
.SYNOPSIS
    Provisions SharePoint lists for the Vehicle Booking Request system.
.DESCRIPTION
    Creates 7 lists with all required fields and indexes.
    Safe to re-run: skips lists/fields that already exist.
    Compatible with PS5 (SharePointPnPPowerShellOnline) and PS7 (PnP.PowerShell).
.USAGE
    .\Provision-VehicleBooking.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module bootstrap
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module PnP.PowerShell -ErrorAction Stop
} else {
    Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue -ErrorAction Stop
}

# ===========================================================================
# Connection
# ===========================================================================

$SiteUrl = "https://powerappautomate.sharepoint.com/sites/Testside"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Vehicle Booking -- SharePoint Provisioning Script"         -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Site : $SiteUrl"                                           -ForegroundColor Gray
Write-Host ""
Write-Host "[CONNECT] Connecting to SharePoint..." -ForegroundColor Yellow
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Connect-PnPOnline -Url $SiteUrl -Interactive
} else {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
}
Write-Host "[CONNECT] Connected successfully." -ForegroundColor Green
Write-Host ""

# ===========================================================================
# Helper: Ensure-List
# ===========================================================================
function Ensure-List {
    param([string]$ListName, [string]$Description)
    $existing = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        Write-Host "  [LIST] '$ListName' already exists -- skipping." -ForegroundColor DarkGray
        return
    }
    Write-Host "  [LIST] Creating '$ListName'..." -ForegroundColor Yellow
    New-PnPList -Title $ListName -Template GenericList -OnQuickLaunch | Out-Null
    if ($Description) {
        Set-PnPList -Identity $ListName -Description $Description
    }
    Write-Host "  [LIST] '$ListName' created." -ForegroundColor Green
}

# ===========================================================================
# Helper: Ensure-Field
# ===========================================================================
function Ensure-Field {
    param(
        [string]$ListName,
        [string]$InternalName,
        [string]$DisplayName,
        [string]$Type,
        [string[]]$Choices
    )
    $existing = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        Write-Host "    [FIELD] '$InternalName' already exists -- skipping." -ForegroundColor DarkGray
        return
    }
    Write-Host "    [FIELD] Adding '$InternalName' ($Type)..." -ForegroundColor Yellow
    if ($Type -eq "Boolean") {
        $xml = "<Field Type='Boolean' DisplayName='$DisplayName' Name='$InternalName' StaticName='$InternalName'><Default>0</Default></Field>"
        Add-PnPFieldFromXml -List $ListName -FieldXml $xml | Out-Null
    } elseif ($Type -eq "Choice") {
        Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Choice -Choices $Choices | Out-Null
    } elseif ($Type -eq "Note") {
        Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Note | Out-Null
    } elseif ($Type -eq "Number") {
        Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Number | Out-Null
    } elseif ($Type -eq "DateTime") {
        Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type DateTime | Out-Null
    } else {
        Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Text | Out-Null
    }
    Write-Host "    [FIELD] '$InternalName' added." -ForegroundColor Green
}

# ===========================================================================
# Helper: Ensure-Index
# ===========================================================================
function Ensure-Index {
    param([string]$ListName, [string]$FieldName)
    try {
        $f = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction SilentlyContinue
        if ($null -eq $f) {
            Write-Host "    [INDEX] Field '$FieldName' not found -- skipping index." -ForegroundColor Red
            return
        }
        if ($f.Indexed) {
            Write-Host "    [INDEX] '$FieldName' already indexed -- skipping." -ForegroundColor DarkGray
            return
        }
        Write-Host "    [INDEX] Indexing '$FieldName'..." -ForegroundColor Yellow
        Set-PnPField -List $ListName -Identity $FieldName -Values @{ Indexed = $true }
        Write-Host "    [INDEX] '$FieldName' indexed." -ForegroundColor Green
    } catch {
        Write-Host "    [WARN ] Could not index '$FieldName': $_" -ForegroundColor DarkYellow
    }
}

# ===========================================================================
# 1. VehicleBookingRequest
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  1/7  VehicleBookingRequest"                                -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

Ensure-List -ListName "VehicleBookingRequest" -Description "Vehicle booking requests submitted by employees"

Ensure-Field "VehicleBookingRequest" "RequestCode"          "Request Code"           "Text"
Ensure-Field "VehicleBookingRequest" "RequesterId"          "Requester ID"           "Text"
Ensure-Field "VehicleBookingRequest" "RequesterName"        "Requester Name"         "Text"
Ensure-Field "VehicleBookingRequest" "RequesterEmail"       "Requester Email"        "Text"
Ensure-Field "VehicleBookingRequest" "Department"           "Department"             "Text"
Ensure-Field "VehicleBookingRequest" "PhoneNumber"          "Phone Number"           "Text"
Ensure-Field "VehicleBookingRequest" "PickupLocation"       "Pickup Location"        "Text"
Ensure-Field "VehicleBookingRequest" "DropoffLocation"      "Dropoff Location"       "Text"
Ensure-Field "VehicleBookingRequest" "PickupDateTime"       "Pickup Date Time"       "DateTime"
Ensure-Field "VehicleBookingRequest" "ReturnDateTime"       "Return Date Time"       "DateTime"
Ensure-Field "VehicleBookingRequest" "IsRoundTrip"          "Is Round Trip"          "Boolean"
Ensure-Field "VehicleBookingRequest" "NumberOfPassengers"   "Number Of Passengers"   "Number"
Ensure-Field "VehicleBookingRequest" "Purpose"              "Purpose"                "Note"
Ensure-Field "VehicleBookingRequest" "VehicleType"          "Vehicle Type"           "Choice"  -Choices @("4-seat","7-seat","16-seat","Truck","Other")
Ensure-Field "VehicleBookingRequest" "SpecialRequirement"   "Special Requirement"    "Note"
Ensure-Field "VehicleBookingRequest" "Status"               "Status"                 "Choice"  -Choices @("DRAFT","SUBMITTED","PENDING_MANAGER_APPROVAL","MANAGER_APPROVED","NEED_MORE_INFORMATION","RESUBMITTED","PENDING_TRANSPORT_ASSIGNMENT","VEHICLE_ASSIGNED","CONFIRMED","DRIVER_CONFIRMED","IN_PROGRESS","COMPLETED","REJECTED","REJECTED_NO_VEHICLE","CANCELLED")
Ensure-Field "VehicleBookingRequest" "CurrentApproverId"    "Current Approver ID"    "Text"
Ensure-Field "VehicleBookingRequest" "CurrentOwnerId"       "Current Owner ID"       "Text"
Ensure-Field "VehicleBookingRequest" "AssignedVehicleId"    "Assigned Vehicle ID"    "Number"
Ensure-Field "VehicleBookingRequest" "AssignedVehiclePlate" "Assigned Vehicle Plate" "Text"
Ensure-Field "VehicleBookingRequest" "AssignedDriverId"     "Assigned Driver ID"     "Number"
Ensure-Field "VehicleBookingRequest" "AssignedDriverName"   "Assigned Driver Name"   "Text"
Ensure-Field "VehicleBookingRequest" "AssignedDriverPhone"  "Assigned Driver Phone"  "Text"
Ensure-Field "VehicleBookingRequest" "ActualStartTime"      "Actual Start Time"      "DateTime"
Ensure-Field "VehicleBookingRequest" "ActualEndTime"        "Actual End Time"        "DateTime"
Ensure-Field "VehicleBookingRequest" "StartOdometer"        "Start Odometer"         "Number"
Ensure-Field "VehicleBookingRequest" "EndOdometer"          "End Odometer"           "Number"
Ensure-Field "VehicleBookingRequest" "TotalDistance"        "Total Distance"         "Number"
Ensure-Field "VehicleBookingRequest" "CancelReason"         "Cancel Reason"          "Note"
Ensure-Field "VehicleBookingRequest" "AdminNote"            "Admin Note"             "Note"
Ensure-Field "VehicleBookingRequest" "SubmittedDate"        "Submitted Date"         "DateTime"
Ensure-Field "VehicleBookingRequest" "CompletedDate"        "Completed Date"         "DateTime"
Ensure-Field "VehicleBookingRequest" "IsDeleted"            "Is Deleted"             "Boolean"

Write-Host "  [INDEX] Indexing VehicleBookingRequest fields..." -ForegroundColor Cyan
Ensure-Index "VehicleBookingRequest" "RequestCode"
Ensure-Index "VehicleBookingRequest" "RequesterEmail"
Ensure-Index "VehicleBookingRequest" "Status"
Ensure-Index "VehicleBookingRequest" "PickupDateTime"
Ensure-Index "VehicleBookingRequest" "AssignedVehicleId"
Ensure-Index "VehicleBookingRequest" "AssignedDriverId"
Ensure-Index "VehicleBookingRequest" "IsDeleted"
Write-Host ""

# ===========================================================================
# 2. VehicleBookingHistory
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  2/7  VehicleBookingHistory"                                -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

Ensure-List -ListName "VehicleBookingHistory" -Description "Audit history of status changes for each booking request"

Ensure-Field "VehicleBookingHistory" "BookingRequestID"  "Booking Request ID"  "Number"
Ensure-Field "VehicleBookingHistory" "RequestCode"       "Request Code"        "Text"
Ensure-Field "VehicleBookingHistory" "Action"            "Action"              "Text"
Ensure-Field "VehicleBookingHistory" "OldStatus"         "Old Status"          "Text"
Ensure-Field "VehicleBookingHistory" "NewStatus"         "New Status"          "Text"
Ensure-Field "VehicleBookingHistory" "ActionByUserId"    "Action By User ID"   "Text"
Ensure-Field "VehicleBookingHistory" "ActionByName"      "Action By Name"      "Text"
Ensure-Field "VehicleBookingHistory" "ActionByEmail"     "Action By Email"     "Text"
Ensure-Field "VehicleBookingHistory" "ActionDate"        "Action Date"         "DateTime"
Ensure-Field "VehicleBookingHistory" "Note"              "Note"                "Note"

Write-Host "  [INDEX] Indexing VehicleBookingHistory fields..." -ForegroundColor Cyan
Ensure-Index "VehicleBookingHistory" "BookingRequestID"
Ensure-Index "VehicleBookingHistory" "RequestCode"
Ensure-Index "VehicleBookingHistory" "ActionDate"
Write-Host ""

# ===========================================================================
# 3. VehicleBookingComment
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  3/7  VehicleBookingComment"                                -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

Ensure-List -ListName "VehicleBookingComment" -Description "Comments and messages attached to booking requests"

Ensure-Field "VehicleBookingComment" "BookingRequestID"  "Booking Request ID"  "Number"
Ensure-Field "VehicleBookingComment" "RequestCode"       "Request Code"        "Text"
Ensure-Field "VehicleBookingComment" "Comment"           "Comment"             "Note"
Ensure-Field "VehicleBookingComment" "CreatedByUserId"   "Created By User ID"  "Text"
Ensure-Field "VehicleBookingComment" "CreatedByName"     "Created By Name"     "Text"
Ensure-Field "VehicleBookingComment" "CreatedByEmail"    "Created By Email"    "Text"
Ensure-Field "VehicleBookingComment" "CommentDate"       "Comment Date"        "DateTime"
Ensure-Field "VehicleBookingComment" "IsDeleted"         "Is Deleted"          "Boolean"

Write-Host "  [INDEX] Indexing VehicleBookingComment fields..." -ForegroundColor Cyan
Ensure-Index "VehicleBookingComment" "BookingRequestID"
Ensure-Index "VehicleBookingComment" "RequestCode"
Ensure-Index "VehicleBookingComment" "CommentDate"
Ensure-Index "VehicleBookingComment" "IsDeleted"
Write-Host ""

# ===========================================================================
# 4. VehicleMaster
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  4/7  VehicleMaster"                                        -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

Ensure-List -ListName "VehicleMaster" -Description "Master list of all company vehicles"

Ensure-Field "VehicleMaster" "VehicleCode"     "Vehicle Code"     "Text"
Ensure-Field "VehicleMaster" "PlateNumber"     "Plate Number"     "Text"
Ensure-Field "VehicleMaster" "VehicleType"     "Vehicle Type"     "Choice"  -Choices @("4-seat","7-seat","16-seat","Truck","Other")
Ensure-Field "VehicleMaster" "Brand"           "Brand"            "Text"
Ensure-Field "VehicleMaster" "Model"           "Model"            "Text"
Ensure-Field "VehicleMaster" "Capacity"        "Capacity"         "Number"
Ensure-Field "VehicleMaster" "Status"          "Status"           "Choice"  -Choices @("AVAILABLE","IN_USE","MAINTENANCE","INACTIVE")
Ensure-Field "VehicleMaster" "CurrentOdometer" "Current Odometer" "Number"
Ensure-Field "VehicleMaster" "Note"            "Note"             "Note"
Ensure-Field "VehicleMaster" "IsDeleted"       "Is Deleted"       "Boolean"

Write-Host "  [INDEX] Indexing VehicleMaster fields..." -ForegroundColor Cyan
Ensure-Index "VehicleMaster" "VehicleCode"
Ensure-Index "VehicleMaster" "PlateNumber"
Ensure-Index "VehicleMaster" "VehicleType"
Ensure-Index "VehicleMaster" "Status"
Ensure-Index "VehicleMaster" "IsDeleted"
Write-Host ""

# ===========================================================================
# 5. DriverMaster
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  5/7  DriverMaster"                                         -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

Ensure-List -ListName "DriverMaster" -Description "Master list of all company drivers"

Ensure-Field "DriverMaster" "DriverCode"     "Driver Code"    "Text"
Ensure-Field "DriverMaster" "DriverName"     "Driver Name"    "Text"
Ensure-Field "DriverMaster" "DriverPhone"    "Driver Phone"   "Text"
Ensure-Field "DriverMaster" "DriverEmail"    "Driver Email"   "Text"
Ensure-Field "DriverMaster" "Status"         "Status"         "Choice"  -Choices @("AVAILABLE","ON_TRIP","INACTIVE")
Ensure-Field "DriverMaster" "LicenseNumber"  "License Number" "Text"
Ensure-Field "DriverMaster" "Note"           "Note"           "Note"
Ensure-Field "DriverMaster" "IsDeleted"      "Is Deleted"     "Boolean"

Write-Host "  [INDEX] Indexing DriverMaster fields..." -ForegroundColor Cyan
Ensure-Index "DriverMaster" "DriverCode"
Ensure-Index "DriverMaster" "DriverEmail"
Ensure-Index "DriverMaster" "Status"
Ensure-Index "DriverMaster" "IsDeleted"
Write-Host ""

# ===========================================================================
# 6. UserRole
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  6/7  UserRole"                                             -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

Ensure-List -ListName "UserRole" -Description "Maps SharePoint users to application roles"

Ensure-Field "UserRole" "UserId"     "User ID"     "Text"
Ensure-Field "UserRole" "UserName"   "User Name"   "Text"
Ensure-Field "UserRole" "UserEmail"  "User Email"  "Text"
Ensure-Field "UserRole" "Role"       "Role"        "Choice"  -Choices @("Requester","Manager","TransportAdmin","Driver","Admin")
Ensure-Field "UserRole" "Department" "Department"  "Text"
Ensure-Field "UserRole" "IsActive"   "Is Active"   "Boolean"

Write-Host "  [INDEX] Indexing UserRole fields..." -ForegroundColor Cyan
Ensure-Index "UserRole" "UserId"
Ensure-Index "UserRole" "UserEmail"
Ensure-Index "UserRole" "Role"
Ensure-Index "UserRole" "IsActive"
Write-Host ""

# ===========================================================================
# 7. LocationMaster
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  7/7  LocationMaster"                                       -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

Ensure-List -ListName "LocationMaster" -Description "Master list of pickup/dropoff locations"

Ensure-Field "LocationMaster" "LocationCode" "Location Code" "Text"
Ensure-Field "LocationMaster" "LocationName" "Location Name" "Text"
Ensure-Field "LocationMaster" "Address"      "Address"       "Note"
Ensure-Field "LocationMaster" "Province"     "Province"      "Text"
Ensure-Field "LocationMaster" "Status"       "Status"        "Choice"  -Choices @("ACTIVE","INACTIVE")
Ensure-Field "LocationMaster" "IsDeleted"    "Is Deleted"    "Boolean"

Write-Host "  [INDEX] Indexing LocationMaster fields..." -ForegroundColor Cyan
Ensure-Index "LocationMaster" "LocationCode"
Ensure-Index "LocationMaster" "Province"
Ensure-Index "LocationMaster" "Status"
Ensure-Index "LocationMaster" "IsDeleted"
Write-Host ""

# ===========================================================================
# Done
# ===========================================================================
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Provisioning completed"                                     -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor White
Write-Host "  VehicleBookingRequest  : 32 fields, 7 indexes" -ForegroundColor Gray
Write-Host "  VehicleBookingHistory  : 10 fields, 3 indexes" -ForegroundColor Gray
Write-Host "  VehicleBookingComment  :  8 fields, 4 indexes" -ForegroundColor Gray
Write-Host "  VehicleMaster          : 10 fields, 5 indexes" -ForegroundColor Gray
Write-Host "  DriverMaster           :  8 fields, 4 indexes" -ForegroundColor Gray
Write-Host "  UserRole               :  6 fields, 4 indexes" -ForegroundColor Gray
Write-Host "  LocationMaster         :  6 fields, 4 indexes" -ForegroundColor Gray
Write-Host ""

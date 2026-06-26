<#
.SYNOPSIS
    Seeds sample data into all 7 Vehicle Booking SharePoint lists.
.DESCRIPTION
    Inserts realistic Vietnamese sample records.
    Safe to re-run: checks for existing records by key field before inserting.
.USAGE
    .\SeedData-VehicleBooking.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Module bootstrap
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module PnP.PowerShell -ErrorAction Stop
} else {
    Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue -ErrorAction Stop
}

$SiteUrl = "https://powerappautomate.sharepoint.com/sites/Testside"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Vehicle Booking -- Seed Sample Data"                        -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Site : $SiteUrl"                                            -ForegroundColor Gray
Write-Host ""

Write-Host "[CONNECT] Connecting..." -ForegroundColor Yellow
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Connect-PnPOnline -Url $SiteUrl -Interactive
} else {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
}
Write-Host "[CONNECT] Connected." -ForegroundColor Green
Write-Host ""

# ===========================================================================
# Helper
# ===========================================================================
function Add-ItemIfNotExists {
    param(
        [string]$ListName,
        [string]$KeyField,
        [string]$KeyValue,
        [hashtable]$Values
    )
    $existing = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='$KeyField'/><Value Type='Text'>$KeyValue</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "    [SKIP] '$KeyValue' already exists." -ForegroundColor DarkGray
        return
    }
    Add-PnPListItem -List $ListName -Values $Values | Out-Null
    Write-Host "    [ADD ] '$KeyValue' inserted." -ForegroundColor Green
}

# ===========================================================================
# 1. LocationMaster  (seed first -- referenced by requests)
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  1/7  LocationMaster"                                        -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$locations = @(
    @{ LocationCode="LOC001"; LocationName="Head Office - Ha Noi";       Address="18 Tam Trinh, Hai Ba Trung, Ha Noi";              Province="Ha Noi";     Status="ACTIVE";   IsDeleted=$false }
    @{ LocationCode="LOC002"; LocationName="Branch Office - HCM";        Address="123 Nguyen Van Linh, Tan Phong, Quan 7, TP.HCM";  Province="TP.HCM";     Status="ACTIVE";   IsDeleted=$false }
    @{ LocationCode="LOC003"; LocationName="Factory - Binh Duong";       Address="KCN My Phuoc 3, Ben Cat, Binh Duong";             Province="Binh Duong"; Status="ACTIVE";   IsDeleted=$false }
    @{ LocationCode="LOC004"; LocationName="Warehouse - Dong Nai";       Address="KCN Long Thanh, Dong Nai";                        Province="Dong Nai";   Status="ACTIVE";   IsDeleted=$false }
    @{ LocationCode="LOC005"; LocationName="Partner Site - Da Nang";     Address="50 Bach Dang, Hai Chau, Da Nang";                 Province="Da Nang";    Status="ACTIVE";   IsDeleted=$false }
    @{ LocationCode="LOC006"; LocationName="Airport - Noi Bai";          Address="San bay Quoc te Noi Bai, Soc Son, Ha Noi";        Province="Ha Noi";     Status="ACTIVE";   IsDeleted=$false }
    @{ LocationCode="LOC007"; LocationName="Airport - Tan Son Nhat";     Address="San bay Tan Son Nhat, Tan Binh, TP.HCM";          Province="TP.HCM";     Status="ACTIVE";   IsDeleted=$false }
    @{ LocationCode="LOC008"; LocationName="Depot - Can Tho";            Address="KCN Tra Noc, Binh Thuy, Can Tho";                 Province="Can Tho";    Status="INACTIVE"; IsDeleted=$false }
)

foreach ($loc in $locations) {
    Add-ItemIfNotExists -ListName "LocationMaster" -KeyField "LocationCode" -KeyValue $loc.LocationCode -Values @{
        Title        = $loc.LocationName
        LocationCode = $loc.LocationCode
        LocationName = $loc.LocationName
        Address      = $loc.Address
        Province     = $loc.Province
        Status       = $loc.Status
        IsDeleted    = $loc.IsDeleted
    }
}
Write-Host ""

# ===========================================================================
# 2. VehicleMaster
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  2/7  VehicleMaster"                                        -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$vehicles = @(
    @{ VehicleCode="VH001"; PlateNumber="30A-12345"; VehicleType="4-seat";  Brand="Toyota";   Model="Vios";      Capacity=4;  Status="AVAILABLE";   CurrentOdometer=45000; Note="Xe moi bao duong thang 5/2026"; IsDeleted=$false }
    @{ VehicleCode="VH002"; PlateNumber="51F-56789"; VehicleType="7-seat";  Brand="Toyota";   Model="Innova";    Capacity=7;  Status="AVAILABLE";   CurrentOdometer=82000; Note="";                               IsDeleted=$false }
    @{ VehicleCode="VH003"; PlateNumber="30K-99001"; VehicleType="16-seat"; Brand="Ford";     Model="Transit";   Capacity=16; Status="IN_USE";       CurrentOdometer=120000;Note="Dang cho chuyen cong tac";       IsDeleted=$false }
    @{ VehicleCode="VH004"; PlateNumber="51G-44321"; VehicleType="7-seat";  Brand="Mitsubishi";Model="Xpander"; Capacity=7;  Status="MAINTENANCE"; CurrentOdometer=67000; Note="Sua hop so, xong 30/06/2026";     IsDeleted=$false }
    @{ VehicleCode="VH005"; PlateNumber="30H-77654"; VehicleType="4-seat";  Brand="Honda";    Model="City";      Capacity=4;  Status="AVAILABLE";   CurrentOdometer=31000; Note="";                               IsDeleted=$false }
    @{ VehicleCode="VH006"; PlateNumber="92C-11223"; VehicleType="Truck";   Brand="Isuzu";    Model="NPR400";    Capacity=2;  Status="AVAILABLE";   CurrentOdometer=95000; Note="Tai trong 3.5 tan";               IsDeleted=$false }
    @{ VehicleCode="VH007"; PlateNumber="51H-33445"; VehicleType="16-seat"; Brand="Hyundai";  Model="Solati";    Capacity=16; Status="AVAILABLE";   CurrentOdometer=54000; Note="";                               IsDeleted=$false }
    @{ VehicleCode="VH008"; PlateNumber="30L-88990"; VehicleType="4-seat";  Brand="Kia";      Model="K3";        Capacity=4;  Status="INACTIVE";    CurrentOdometer=210000;Note="Xe cu, cho thanh ly";              IsDeleted=$false }
)

foreach ($v in $vehicles) {
    Add-ItemIfNotExists -ListName "VehicleMaster" -KeyField "VehicleCode" -KeyValue $v.VehicleCode -Values @{
        Title            = "$($v.PlateNumber) - $($v.Brand) $($v.Model)"
        VehicleCode      = $v.VehicleCode
        PlateNumber      = $v.PlateNumber
        VehicleType      = $v.VehicleType
        Brand            = $v.Brand
        Model            = $v.Model
        Capacity         = $v.Capacity
        Status           = $v.Status
        CurrentOdometer  = $v.CurrentOdometer
        Note             = $v.Note
        IsDeleted        = $v.IsDeleted
    }
}
Write-Host ""

# ===========================================================================
# 3. DriverMaster
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  3/7  DriverMaster"                                         -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$drivers = @(
    @{ DriverCode="DRV001"; DriverName="Nguyen Van An";    DriverPhone="0912345601"; DriverEmail="an.nv@greenfeed.com.vn";    Status="AVAILABLE"; LicenseNumber="B2-012345"; Note="Lai xe kinh nghiem 10 nam"; IsDeleted=$false }
    @{ DriverCode="DRV002"; DriverName="Tran Minh Duc";    DriverPhone="0923456702"; DriverEmail="duc.tm@greenfeed.com.vn";   Status="ON_TRIP";   LicenseNumber="B2-023456"; Note="";                         IsDeleted=$false }
    @{ DriverCode="DRV003"; DriverName="Le Thanh Hung";    DriverPhone="0934567803"; DriverEmail="hung.lt@greenfeed.com.vn";  Status="AVAILABLE"; LicenseNumber="D-034567";  Note="Lai xe tai, xe khach";     IsDeleted=$false }
    @{ DriverCode="DRV004"; DriverName="Pham Quoc Bao";    DriverPhone="0945678904"; DriverEmail="bao.pq@greenfeed.com.vn";   Status="AVAILABLE"; LicenseNumber="B2-045678"; Note="";                         IsDeleted=$false }
    @{ DriverCode="DRV005"; DriverName="Hoang Van Cuong";  DriverPhone="0956789005"; DriverEmail="cuong.hv@greenfeed.com.vn"; Status="INACTIVE";  LicenseNumber="B2-056789"; Note="Nghi viec tu 01/06/2026";  IsDeleted=$false }
    @{ DriverCode="DRV006"; DriverName="Vo Thi Mai";       DriverPhone="0967890106"; DriverEmail="mai.vt@greenfeed.com.vn";   Status="AVAILABLE"; LicenseNumber="B2-067890"; Note="";                         IsDeleted=$false }
)

foreach ($d in $drivers) {
    Add-ItemIfNotExists -ListName "DriverMaster" -KeyField "DriverCode" -KeyValue $d.DriverCode -Values @{
        Title         = $d.DriverName
        DriverCode    = $d.DriverCode
        DriverName    = $d.DriverName
        DriverPhone   = $d.DriverPhone
        DriverEmail   = $d.DriverEmail
        Status        = $d.Status
        LicenseNumber = $d.LicenseNumber
        Note          = $d.Note
        IsDeleted     = $d.IsDeleted
    }
}
Write-Host ""

# ===========================================================================
# 4. UserRole
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  4/7  UserRole"                                             -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$users = @(
    @{ UserId="U001"; UserName="Nguyen Thi Bich";    UserEmail="bich.nt@greenfeed.com.vn";    Role="Requester";      Department="Sales";     IsActive=$true  }
    @{ UserId="U002"; UserName="Tran Van Khanh";     UserEmail="khanh.tv@greenfeed.com.vn";   Role="Requester";      Department="Marketing"; IsActive=$true  }
    @{ UserId="U003"; UserName="Le Thi Huong";       UserEmail="huong.lt@greenfeed.com.vn";   Role="Manager";        Department="Sales";     IsActive=$true  }
    @{ UserId="U004"; UserName="Pham Thanh Long";    UserEmail="long.pt@greenfeed.com.vn";    Role="Manager";        Department="Marketing"; IsActive=$true  }
    @{ UserId="U005"; UserName="Hoang Minh Tu";      UserEmail="tu.hm@greenfeed.com.vn";      Role="TransportAdmin"; Department="Admin";     IsActive=$true  }
    @{ UserId="U006"; UserName="Nguyen Van An";      UserEmail="an.nv@greenfeed.com.vn";      Role="Driver";         Department="Transport"; IsActive=$true  }
    @{ UserId="U007"; UserName="Tran Minh Duc";      UserEmail="duc.tm@greenfeed.com.vn";     Role="Driver";         Department="Transport"; IsActive=$true  }
    @{ UserId="U008"; UserName="Le Thanh Hung";      UserEmail="hung.lt@greenfeed.com.vn";    Role="Driver";         Department="Transport"; IsActive=$true  }
    @{ UserId="U009"; UserName="Admin System";       UserEmail="admin@greenfeed.com.vn";      Role="Admin";          Department="IT";        IsActive=$true  }
    @{ UserId="U010"; UserName="Vo Thi Lan";         UserEmail="lan.vt@greenfeed.com.vn";     Role="Requester";      Department="HR";        IsActive=$false }
)

foreach ($u in $users) {
    Add-ItemIfNotExists -ListName "UserRole" -KeyField "UserId" -KeyValue $u.UserId -Values @{
        Title      = $u.UserName
        UserId     = $u.UserId
        UserName   = $u.UserName
        UserEmail  = $u.UserEmail
        Role       = $u.Role
        Department = $u.Department
        IsActive   = $u.IsActive
    }
}
Write-Host ""

# ===========================================================================
# 5. VehicleBookingRequest
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  5/7  VehicleBookingRequest"                                -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$requests = @(
    @{
        RequestCode         = "VBR-2026-001"
        RequesterId         = "U001"
        RequesterName       = "Nguyen Thi Bich"
        RequesterEmail      = "bich.nt@greenfeed.com.vn"
        Department          = "Sales"
        PhoneNumber         = "0912111001"
        PickupLocation      = "Head Office - Ha Noi"
        DropoffLocation     = "Partner Site - Da Nang"
        PickupDateTime      = "2026-07-01T08:00:00"
        ReturnDateTime      = "2026-07-03T18:00:00"
        IsRoundTrip         = $true
        NumberOfPassengers  = 3
        Purpose             = "Chuyen cong tac gap doi tac tai Da Nang, ky ket hop dong Q3/2026"
        VehicleType         = "7-seat"
        SpecialRequirement  = ""
        Status              = "COMPLETED"
        CurrentApproverId   = "U003"
        CurrentOwnerId      = "U005"
        AssignedVehicleId   = 2
        AssignedVehiclePlate= "51F-56789"
        AssignedDriverId    = 1
        AssignedDriverName  = "Nguyen Van An"
        AssignedDriverPhone = "0912345601"
        ActualStartTime     = "2026-07-01T08:15:00"
        ActualEndTime       = "2026-07-03T19:30:00"
        StartOdometer       = 80000
        EndOdometer         = 81240
        TotalDistance       = 1240
        CancelReason        = ""
        AdminNote           = "Hoan thanh tot"
        SubmittedDate       = "2026-06-25T09:00:00"
        CompletedDate       = "2026-07-03T20:00:00"
        IsDeleted           = $false
    }
    @{
        RequestCode         = "VBR-2026-002"
        RequesterId         = "U002"
        RequesterName       = "Tran Van Khanh"
        RequesterEmail      = "khanh.tv@greenfeed.com.vn"
        Department          = "Marketing"
        PhoneNumber         = "0923222002"
        PickupLocation      = "Branch Office - HCM"
        DropoffLocation     = "Factory - Binh Duong"
        PickupDateTime      = "2026-07-05T07:30:00"
        ReturnDateTime      = "2026-07-05T17:00:00"
        IsRoundTrip         = $true
        NumberOfPassengers  = 5
        Purpose             = "Tham quan nha may, chup anh san pham cho chien dich marketing"
        VehicleType         = "7-seat"
        SpecialRequirement  = "Can xe co dieu hoa"
        Status              = "DRIVER_CONFIRMED"
        CurrentApproverId   = "U004"
        CurrentOwnerId      = "U005"
        AssignedVehicleId   = 2
        AssignedVehiclePlate= "51F-56789"
        AssignedDriverId    = 2
        AssignedDriverName  = "Tran Minh Duc"
        AssignedDriverPhone = "0923456702"
        ActualStartTime     = ""
        ActualEndTime       = ""
        StartOdometer       = 0
        EndOdometer         = 0
        TotalDistance       = 0
        CancelReason        = ""
        AdminNote           = ""
        SubmittedDate       = "2026-06-26T10:30:00"
        CompletedDate       = ""
        IsDeleted           = $false
    }
    @{
        RequestCode         = "VBR-2026-003"
        RequesterId         = "U001"
        RequesterName       = "Nguyen Thi Bich"
        RequesterEmail      = "bich.nt@greenfeed.com.vn"
        Department          = "Sales"
        PhoneNumber         = "0912111001"
        PickupLocation      = "Head Office - Ha Noi"
        DropoffLocation     = "Airport - Noi Bai"
        PickupDateTime      = "2026-07-10T05:30:00"
        ReturnDateTime      = ""
        IsRoundTrip         = $false
        NumberOfPassengers  = 1
        Purpose             = "Ra san bay don giam doc dieu hanh tu nuoc ngoai ve"
        VehicleType         = "4-seat"
        SpecialRequirement  = "Xe sach, lay som"
        Status              = "PENDING_MANAGER_APPROVAL"
        CurrentApproverId   = "U003"
        CurrentOwnerId      = ""
        AssignedVehicleId   = 0
        AssignedVehiclePlate= ""
        AssignedDriverId    = 0
        AssignedDriverName  = ""
        AssignedDriverPhone = ""
        ActualStartTime     = ""
        ActualEndTime       = ""
        StartOdometer       = 0
        EndOdometer         = 0
        TotalDistance       = 0
        CancelReason        = ""
        AdminNote           = ""
        SubmittedDate       = "2026-06-26T14:00:00"
        CompletedDate       = ""
        IsDeleted           = $false
    }
    @{
        RequestCode         = "VBR-2026-004"
        RequesterId         = "U010"
        RequesterName       = "Vo Thi Lan"
        RequesterEmail      = "lan.vt@greenfeed.com.vn"
        Department          = "HR"
        PhoneNumber         = "0956666010"
        PickupLocation      = "Head Office - Ha Noi"
        DropoffLocation     = "Branch Office - HCM"
        PickupDateTime      = "2026-06-20T06:00:00"
        ReturnDateTime      = "2026-06-22T20:00:00"
        IsRoundTrip         = $true
        NumberOfPassengers  = 2
        Purpose             = "Di cong tac tuyen dung chi nhanh HCM"
        VehicleType         = "4-seat"
        SpecialRequirement  = ""
        Status              = "REJECTED"
        CurrentApproverId   = "U003"
        CurrentOwnerId      = ""
        AssignedVehicleId   = 0
        AssignedVehiclePlate= ""
        AssignedDriverId    = 0
        AssignedDriverName  = ""
        AssignedDriverPhone = ""
        ActualStartTime     = ""
        ActualEndTime       = ""
        StartOdometer       = 0
        EndOdometer         = 0
        TotalDistance       = 0
        CancelReason        = ""
        AdminNote           = "Khong co xe kha dung trong khoang thoi gian nay"
        SubmittedDate       = "2026-06-15T11:00:00"
        CompletedDate       = ""
        IsDeleted           = $false
    }
    @{
        RequestCode         = "VBR-2026-005"
        RequesterId         = "U002"
        RequesterName       = "Tran Van Khanh"
        RequesterEmail      = "khanh.tv@greenfeed.com.vn"
        Department          = "Marketing"
        PhoneNumber         = "0923222002"
        PickupLocation      = "Branch Office - HCM"
        DropoffLocation     = "Warehouse - Dong Nai"
        PickupDateTime      = "2026-07-08T08:00:00"
        ReturnDateTime      = "2026-07-08T16:00:00"
        IsRoundTrip         = $true
        NumberOfPassengers  = 8
        Purpose             = "Tham quan kho hang phuc vu bai viet noi dung website"
        VehicleType         = "16-seat"
        SpecialRequirement  = "Can xe 16 cho vi co nhieu nguoi"
        Status              = "VEHICLE_ASSIGNED"
        CurrentApproverId   = "U004"
        CurrentOwnerId      = "U005"
        AssignedVehicleId   = 7
        AssignedVehiclePlate= "51H-33445"
        AssignedDriverId    = 3
        AssignedDriverName  = "Le Thanh Hung"
        AssignedDriverPhone = "0934567803"
        ActualStartTime     = ""
        ActualEndTime       = ""
        StartOdometer       = 0
        EndOdometer         = 0
        TotalDistance       = 0
        CancelReason        = ""
        AdminNote           = "Da phan cong xe Hyundai Solati"
        SubmittedDate       = "2026-06-26T08:00:00"
        CompletedDate       = ""
        IsDeleted           = $false
    }
    @{
        RequestCode         = "VBR-2026-006"
        RequesterId         = "U001"
        RequesterName       = "Nguyen Thi Bich"
        RequesterEmail      = "bich.nt@greenfeed.com.vn"
        Department          = "Sales"
        PhoneNumber         = "0912111001"
        PickupLocation      = "Head Office - Ha Noi"
        DropoffLocation     = "Partner Site - Da Nang"
        PickupDateTime      = "2026-08-01T07:00:00"
        ReturnDateTime      = "2026-08-02T20:00:00"
        IsRoundTrip         = $true
        NumberOfPassengers  = 4
        Purpose             = "Trien khai chuong trinh khuyen mai mua thu"
        VehicleType         = "7-seat"
        SpecialRequirement  = ""
        Status              = "DRAFT"
        CurrentApproverId   = ""
        CurrentOwnerId      = ""
        AssignedVehicleId   = 0
        AssignedVehiclePlate= ""
        AssignedDriverId    = 0
        AssignedDriverName  = ""
        AssignedDriverPhone = ""
        ActualStartTime     = ""
        ActualEndTime       = ""
        StartOdometer       = 0
        EndOdometer         = 0
        TotalDistance       = 0
        CancelReason        = ""
        AdminNote           = ""
        SubmittedDate       = ""
        CompletedDate       = ""
        IsDeleted           = $false
    }
    @{
        RequestCode         = "VBR-2026-007"
        RequesterId         = "U002"
        RequesterName       = "Tran Van Khanh"
        RequesterEmail      = "khanh.tv@greenfeed.com.vn"
        Department          = "Marketing"
        PhoneNumber         = "0923222002"
        PickupLocation      = "Branch Office - HCM"
        DropoffLocation     = "Airport - Tan Son Nhat"
        PickupDateTime      = "2026-06-18T04:30:00"
        ReturnDateTime      = ""
        IsRoundTrip         = $false
        NumberOfPassengers  = 3
        Purpose             = "Dua doan khach hang nuoc ngoai ra san bay"
        VehicleType         = "7-seat"
        SpecialRequirement  = "Xe sach, tai xe mac ao so mi"
        Status              = "CANCELLED"
        CurrentApproverId   = "U004"
        CurrentOwnerId      = "U005"
        AssignedVehicleId   = 0
        AssignedVehiclePlate= ""
        AssignedDriverId    = 0
        AssignedDriverName  = ""
        AssignedDriverPhone = ""
        ActualStartTime     = ""
        ActualEndTime       = ""
        StartOdometer       = 0
        EndOdometer         = 0
        TotalDistance       = 0
        CancelReason        = "Khach hang doi lich chuyen bay, huy chuyen"
        AdminNote           = ""
        SubmittedDate       = "2026-06-10T09:00:00"
        CompletedDate       = ""
        IsDeleted           = $false
    }
    @{
        RequestCode         = "VBR-2026-008"
        RequesterId         = "U001"
        RequesterName       = "Nguyen Thi Bich"
        RequesterEmail      = "bich.nt@greenfeed.com.vn"
        Department          = "Sales"
        PhoneNumber         = "0912111001"
        PickupLocation      = "Head Office - Ha Noi"
        DropoffLocation     = "Factory - Binh Duong"
        PickupDateTime      = "2026-07-15T08:00:00"
        ReturnDateTime      = "2026-07-15T18:00:00"
        IsRoundTrip         = $true
        NumberOfPassengers  = 6
        Purpose             = "Kiem tra chat luong san pham truoc khi xuat hang"
        VehicleType         = "7-seat"
        SpecialRequirement  = ""
        Status              = "NEED_MORE_INFORMATION"
        CurrentApproverId   = "U003"
        CurrentOwnerId      = ""
        AssignedVehicleId   = 0
        AssignedVehiclePlate= ""
        AssignedDriverId    = 0
        AssignedDriverName  = ""
        AssignedDriverPhone = ""
        ActualStartTime     = ""
        ActualEndTime       = ""
        StartOdometer       = 0
        EndOdometer         = 0
        TotalDistance       = 0
        CancelReason        = ""
        AdminNote           = "Can biet them so luong chinh xac va muc dich cu the"
        SubmittedDate       = "2026-06-26T15:00:00"
        CompletedDate       = ""
        IsDeleted           = $false
    }
)

foreach ($r in $requests) {
    $vals = @{
        Title               = $r.RequestCode
        RequestCode         = $r.RequestCode
        RequesterId         = $r.RequesterId
        RequesterName       = $r.RequesterName
        RequesterEmail      = $r.RequesterEmail
        Department          = $r.Department
        PhoneNumber         = $r.PhoneNumber
        PickupLocation      = $r.PickupLocation
        DropoffLocation     = $r.DropoffLocation
        IsRoundTrip         = $r.IsRoundTrip
        NumberOfPassengers  = $r.NumberOfPassengers
        Purpose             = $r.Purpose
        VehicleType         = $r.VehicleType
        SpecialRequirement  = $r.SpecialRequirement
        Status              = $r.Status
        CurrentApproverId   = $r.CurrentApproverId
        CurrentOwnerId      = $r.CurrentOwnerId
        AssignedVehiclePlate= $r.AssignedVehiclePlate
        AssignedDriverName  = $r.AssignedDriverName
        AssignedDriverPhone = $r.AssignedDriverPhone
        CancelReason        = $r.CancelReason
        AdminNote           = $r.AdminNote
        IsDeleted           = $r.IsDeleted
    }
    if ($r.AssignedVehicleId -gt 0) { $vals["AssignedVehicleId"] = $r.AssignedVehicleId }
    if ($r.AssignedDriverId  -gt 0) { $vals["AssignedDriverId"]  = $r.AssignedDriverId  }
    if ($r.NumberOfPassengers -gt 0){ $vals["NumberOfPassengers"] = $r.NumberOfPassengers }
    if ($r.StartOdometer -gt 0)     { $vals["StartOdometer"]     = $r.StartOdometer      }
    if ($r.EndOdometer   -gt 0)     { $vals["EndOdometer"]       = $r.EndOdometer        }
    if ($r.TotalDistance -gt 0)     { $vals["TotalDistance"]     = $r.TotalDistance      }
    if ($r.PickupDateTime)          { $vals["PickupDateTime"]    = [DateTime]$r.PickupDateTime }
    if ($r.ReturnDateTime)          { $vals["ReturnDateTime"]    = [DateTime]$r.ReturnDateTime }
    if ($r.ActualStartTime)         { $vals["ActualStartTime"]   = [DateTime]$r.ActualStartTime }
    if ($r.ActualEndTime)           { $vals["ActualEndTime"]     = [DateTime]$r.ActualEndTime   }
    if ($r.SubmittedDate)           { $vals["SubmittedDate"]     = [DateTime]$r.SubmittedDate   }
    if ($r.CompletedDate)           { $vals["CompletedDate"]     = [DateTime]$r.CompletedDate   }

    Add-ItemIfNotExists -ListName "VehicleBookingRequest" -KeyField "RequestCode" -KeyValue $r.RequestCode -Values $vals
}
Write-Host ""

# ===========================================================================
# 6. VehicleBookingHistory
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  6/7  VehicleBookingHistory"                                -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

# History uses Number key (BookingRequestID) -- use Title as dedup key
$histories = @(
    # VBR-2026-001 full lifecycle
    @{ Title="VBR-2026-001-H1"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="SUBMIT";           OldStatus="DRAFT";                     NewStatus="SUBMITTED";                    ActionByUserId="U001"; ActionByName="Nguyen Thi Bich"; ActionByEmail="bich.nt@greenfeed.com.vn";  ActionDate="2026-06-25T09:00:00"; Note="Gui yeu cau" }
    @{ Title="VBR-2026-001-H2"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="APPROVE_MANAGER";  OldStatus="SUBMITTED";                 NewStatus="PENDING_MANAGER_APPROVAL";     ActionByUserId="U003"; ActionByName="Le Thi Huong";    ActionByEmail="huong.lt@greenfeed.com.vn"; ActionDate="2026-06-25T11:00:00"; Note="Duyet cap truong phong" }
    @{ Title="VBR-2026-001-H3"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="MANAGER_APPROVED"; OldStatus="PENDING_MANAGER_APPROVAL";  NewStatus="PENDING_TRANSPORT_ASSIGNMENT"; ActionByUserId="U003"; ActionByName="Le Thi Huong";    ActionByEmail="huong.lt@greenfeed.com.vn"; ActionDate="2026-06-25T14:00:00"; Note="Approved" }
    @{ Title="VBR-2026-001-H4"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="ASSIGN_VEHICLE";   OldStatus="PENDING_TRANSPORT_ASSIGNMENT";NewStatus="VEHICLE_ASSIGNED";           ActionByUserId="U005"; ActionByName="Hoang Minh Tu";   ActionByEmail="tu.hm@greenfeed.com.vn";    ActionDate="2026-06-26T08:00:00"; Note="Phan cong xe 51F-56789" }
    @{ Title="VBR-2026-001-H5"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="CONFIRM";          OldStatus="VEHICLE_ASSIGNED";          NewStatus="CONFIRMED";                    ActionByUserId="U005"; ActionByName="Hoang Minh Tu";   ActionByEmail="tu.hm@greenfeed.com.vn";    ActionDate="2026-06-26T09:00:00"; Note="Xac nhan lich" }
    @{ Title="VBR-2026-001-H6"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="DRIVER_CONFIRM";   OldStatus="CONFIRMED";                 NewStatus="DRIVER_CONFIRMED";             ActionByUserId="U006"; ActionByName="Nguyen Van An";   ActionByEmail="an.nv@greenfeed.com.vn";    ActionDate="2026-06-26T10:00:00"; Note="Tai xe xac nhan nhan chuyen" }
    @{ Title="VBR-2026-001-H7"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="START_TRIP";       OldStatus="DRIVER_CONFIRMED";          NewStatus="IN_PROGRESS";                  ActionByUserId="U006"; ActionByName="Nguyen Van An";   ActionByEmail="an.nv@greenfeed.com.vn";    ActionDate="2026-07-01T08:15:00"; Note="Bat dau chuyen di, ODO: 80000" }
    @{ Title="VBR-2026-001-H8"; BookingRequestID=1; RequestCode="VBR-2026-001"; Action="COMPLETE";         OldStatus="IN_PROGRESS";               NewStatus="COMPLETED";                    ActionByUserId="U005"; ActionByName="Hoang Minh Tu";   ActionByEmail="tu.hm@greenfeed.com.vn";    ActionDate="2026-07-03T20:00:00"; Note="Hoan thanh, ODO: 81240, km: 1240" }
    # VBR-2026-002
    @{ Title="VBR-2026-002-H1"; BookingRequestID=2; RequestCode="VBR-2026-002"; Action="SUBMIT";           OldStatus="DRAFT";                     NewStatus="SUBMITTED";                    ActionByUserId="U002"; ActionByName="Tran Van Khanh";  ActionByEmail="khanh.tv@greenfeed.com.vn"; ActionDate="2026-06-26T10:30:00"; Note="" }
    @{ Title="VBR-2026-002-H2"; BookingRequestID=2; RequestCode="VBR-2026-002"; Action="MANAGER_APPROVED"; OldStatus="SUBMITTED";                 NewStatus="PENDING_TRANSPORT_ASSIGNMENT"; ActionByUserId="U004"; ActionByName="Pham Thanh Long"; ActionByEmail="long.pt@greenfeed.com.vn";  ActionDate="2026-06-26T13:00:00"; Note="OK" }
    @{ Title="VBR-2026-002-H3"; BookingRequestID=2; RequestCode="VBR-2026-002"; Action="ASSIGN_VEHICLE";   OldStatus="PENDING_TRANSPORT_ASSIGNMENT";NewStatus="VEHICLE_ASSIGNED";           ActionByUserId="U005"; ActionByName="Hoang Minh Tu";   ActionByEmail="tu.hm@greenfeed.com.vn";    ActionDate="2026-06-26T15:00:00"; Note="Phan cong xe 51F-56789 + tai xe Duc" }
    @{ Title="VBR-2026-002-H4"; BookingRequestID=2; RequestCode="VBR-2026-002"; Action="DRIVER_CONFIRM";   OldStatus="VEHICLE_ASSIGNED";          NewStatus="DRIVER_CONFIRMED";             ActionByUserId="U007"; ActionByName="Tran Minh Duc";   ActionByEmail="duc.tm@greenfeed.com.vn";    ActionDate="2026-06-26T16:30:00"; Note="Da nhan lich" }
    # VBR-2026-003
    @{ Title="VBR-2026-003-H1"; BookingRequestID=3; RequestCode="VBR-2026-003"; Action="SUBMIT";           OldStatus="DRAFT";                     NewStatus="SUBMITTED";                    ActionByUserId="U001"; ActionByName="Nguyen Thi Bich"; ActionByEmail="bich.nt@greenfeed.com.vn";  ActionDate="2026-06-26T14:00:00"; Note="" }
    @{ Title="VBR-2026-003-H2"; BookingRequestID=3; RequestCode="VBR-2026-003"; Action="PENDING_APPROVE";  OldStatus="SUBMITTED";                 NewStatus="PENDING_MANAGER_APPROVAL";     ActionByUserId="U009"; ActionByName="Admin System";    ActionByEmail="admin@greenfeed.com.vn";     ActionDate="2026-06-26T14:05:00"; Note="Chuyen cho truong phong duyet" }
    # VBR-2026-004
    @{ Title="VBR-2026-004-H1"; BookingRequestID=4; RequestCode="VBR-2026-004"; Action="SUBMIT";           OldStatus="DRAFT";                     NewStatus="SUBMITTED";                    ActionByUserId="U010"; ActionByName="Vo Thi Lan";      ActionByEmail="lan.vt@greenfeed.com.vn";    ActionDate="2026-06-15T11:00:00"; Note="" }
    @{ Title="VBR-2026-004-H2"; BookingRequestID=4; RequestCode="VBR-2026-004"; Action="REJECT";           OldStatus="SUBMITTED";                 NewStatus="REJECTED";                     ActionByUserId="U003"; ActionByName="Le Thi Huong";    ActionByEmail="huong.lt@greenfeed.com.vn"; ActionDate="2026-06-16T09:00:00"; Note="Khong co xe trong thoi gian nay" }
    # VBR-2026-007
    @{ Title="VBR-2026-007-H1"; BookingRequestID=7; RequestCode="VBR-2026-007"; Action="SUBMIT";           OldStatus="DRAFT";                     NewStatus="SUBMITTED";                    ActionByUserId="U002"; ActionByName="Tran Van Khanh";  ActionByEmail="khanh.tv@greenfeed.com.vn"; ActionDate="2026-06-10T09:00:00"; Note="" }
    @{ Title="VBR-2026-007-H2"; BookingRequestID=7; RequestCode="VBR-2026-007"; Action="MANAGER_APPROVED"; OldStatus="SUBMITTED";                 NewStatus="PENDING_TRANSPORT_ASSIGNMENT"; ActionByUserId="U004"; ActionByName="Pham Thanh Long"; ActionByEmail="long.pt@greenfeed.com.vn";  ActionDate="2026-06-10T14:00:00"; Note="Duyet" }
    @{ Title="VBR-2026-007-H3"; BookingRequestID=7; RequestCode="VBR-2026-007"; Action="CANCEL";           OldStatus="PENDING_TRANSPORT_ASSIGNMENT";NewStatus="CANCELLED";                  ActionByUserId="U002"; ActionByName="Tran Van Khanh";  ActionByEmail="khanh.tv@greenfeed.com.vn"; ActionDate="2026-06-12T08:00:00"; Note="Khach doi lich chuyen bay" }
)

foreach ($h in $histories) {
    $existing = Get-PnPListItem -List "VehicleBookingHistory" `
        -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($h.Title)</Value></Eq></Where></Query></View>" `
        -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "    [SKIP] '$($h.Title)' already exists." -ForegroundColor DarkGray
        continue
    }
    Add-PnPListItem -List "VehicleBookingHistory" -Values @{
        Title            = $h.Title
        BookingRequestID = $h.BookingRequestID
        RequestCode      = $h.RequestCode
        Action           = $h.Action
        OldStatus        = $h.OldStatus
        NewStatus        = $h.NewStatus
        ActionByUserId   = $h.ActionByUserId
        ActionByName     = $h.ActionByName
        ActionByEmail    = $h.ActionByEmail
        ActionDate       = [DateTime]$h.ActionDate
        Note             = $h.Note
    } | Out-Null
    Write-Host "    [ADD ] '$($h.Title)' inserted." -ForegroundColor Green
}
Write-Host ""

# ===========================================================================
# 7. VehicleBookingComment
# ===========================================================================
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  7/7  VehicleBookingComment"                                -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$comments = @(
    @{ Title="VBR-2026-001-C1"; BookingRequestID=1; RequestCode="VBR-2026-001"; Comment="Anh/chi vui long xac nhan so hanh ly mang theo de bo tri xe cho phu hop.";        CreatedByUserId="U005"; CreatedByName="Hoang Minh Tu";   CreatedByEmail="tu.hm@greenfeed.com.vn";    CommentDate="2026-06-25T15:00:00"; IsDeleted=$false }
    @{ Title="VBR-2026-001-C2"; BookingRequestID=1; RequestCode="VBR-2026-001"; Comment="Chung toi mang theo 3 vali nho, khong can them khong gian hanh ly.";               CreatedByUserId="U001"; CreatedByName="Nguyen Thi Bich"; CreatedByEmail="bich.nt@greenfeed.com.vn";  CommentDate="2026-06-25T16:30:00"; IsDeleted=$false }
    @{ Title="VBR-2026-001-C3"; BookingRequestID=1; RequestCode="VBR-2026-001"; Comment="Da ghi nhan. Tai xe Nguyen Van An se co mat luc 07:45 tai toa nha.";               CreatedByUserId="U005"; CreatedByName="Hoang Minh Tu";   CreatedByEmail="tu.hm@greenfeed.com.vn";    CommentDate="2026-06-26T10:00:00"; IsDeleted=$false }
    @{ Title="VBR-2026-002-C1"; BookingRequestID=2; RequestCode="VBR-2026-002"; Comment="Nho mang theo giay to cong tac cho ca doan.";                                       CreatedByUserId="U004"; CreatedByName="Pham Thanh Long"; CreatedByEmail="long.pt@greenfeed.com.vn";  CommentDate="2026-06-26T13:30:00"; IsDeleted=$false }
    @{ Title="VBR-2026-002-C2"; BookingRequestID=2; RequestCode="VBR-2026-002"; Comment="Da chuan bi du. Cam on anh.";                                                      CreatedByUserId="U002"; CreatedByName="Tran Van Khanh";  CreatedByEmail="khanh.tv@greenfeed.com.vn"; CommentDate="2026-06-26T14:00:00"; IsDeleted=$false }
    @{ Title="VBR-2026-003-C1"; BookingRequestID=3; RequestCode="VBR-2026-003"; Comment="Yeu cau kha gap, de nghi truong phong duyet som trong hom nay.";                   CreatedByUserId="U001"; CreatedByName="Nguyen Thi Bich"; CreatedByEmail="bich.nt@greenfeed.com.vn";  CommentDate="2026-06-26T14:10:00"; IsDeleted=$false }
    @{ Title="VBR-2026-004-C1"; BookingRequestID=4; RequestCode="VBR-2026-004"; Comment="Rat tiec hien tai khong co xe phu hop. Ban co the doi lich sang tuan sau khong?"; CreatedByUserId="U003"; CreatedByName="Le Thi Huong";    CreatedByEmail="huong.lt@greenfeed.com.vn"; CommentDate="2026-06-16T09:05:00"; IsDeleted=$false }
    @{ Title="VBR-2026-004-C2"; BookingRequestID=4; RequestCode="VBR-2026-004"; Comment="Toi se lien he lai voi bo phan de doi lich. Cam on chi.";                          CreatedByUserId="U010"; CreatedByName="Vo Thi Lan";      CreatedByEmail="lan.vt@greenfeed.com.vn";   CommentDate="2026-06-16T10:00:00"; IsDeleted=$false }
    @{ Title="VBR-2026-005-C1"; BookingRequestID=5; RequestCode="VBR-2026-005"; Comment="Da phan cong xe Solati 51H-33445, tai xe Le Thanh Hung se lien lac truoc 1 ngay."; CreatedByUserId="U005"; CreatedByName="Hoang Minh Tu";   CreatedByEmail="tu.hm@greenfeed.com.vn";    CommentDate="2026-06-26T16:00:00"; IsDeleted=$false }
    @{ Title="VBR-2026-008-C1"; BookingRequestID=8; RequestCode="VBR-2026-008"; Comment="Ban vui long cung cap danh sach nguoi tham gia va muc dich kiem tra cu the.";      CreatedByUserId="U003"; CreatedByName="Le Thi Huong";    CreatedByEmail="huong.lt@greenfeed.com.vn"; CommentDate="2026-06-26T15:30:00"; IsDeleted=$false }
    @{ Title="VBR-2026-007-C1"; BookingRequestID=7; RequestCode="VBR-2026-007"; Comment="Khach thay doi chuyen bay, phai huy chuyen. Xin loi vi su bat tien.";              CreatedByUserId="U002"; CreatedByName="Tran Van Khanh";  CreatedByEmail="khanh.tv@greenfeed.com.vn"; CommentDate="2026-06-12T08:05:00"; IsDeleted=$false }
)

foreach ($c in $comments) {
    $existing = Get-PnPListItem -List "VehicleBookingComment" `
        -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($c.Title)</Value></Eq></Where></Query></View>" `
        -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "    [SKIP] '$($c.Title)' already exists." -ForegroundColor DarkGray
        continue
    }
    Add-PnPListItem -List "VehicleBookingComment" -Values @{
        Title            = $c.Title
        BookingRequestID = $c.BookingRequestID
        RequestCode      = $c.RequestCode
        Comment          = $c.Comment
        CreatedByUserId  = $c.CreatedByUserId
        CreatedByName    = $c.CreatedByName
        CreatedByEmail   = $c.CreatedByEmail
        CommentDate      = [DateTime]$c.CommentDate
        IsDeleted        = $c.IsDeleted
    } | Out-Null
    Write-Host "    [ADD ] '$($c.Title)' inserted." -ForegroundColor Green
}
Write-Host ""

# ===========================================================================
# Done
# ===========================================================================
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Seed data completed"                                         -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Records inserted:" -ForegroundColor White
Write-Host "  LocationMaster         :  8 records" -ForegroundColor Gray
Write-Host "  VehicleMaster          :  8 records" -ForegroundColor Gray
Write-Host "  DriverMaster           :  6 records" -ForegroundColor Gray
Write-Host "  UserRole               : 10 records" -ForegroundColor Gray
Write-Host "  VehicleBookingRequest  :  8 records (DRAFT/SUBMITTED/PENDING/APPROVED/ASSIGNED/DRIVER_CONFIRMED/IN_PROGRESS/COMPLETED/REJECTED/CANCELLED)" -ForegroundColor Gray
Write-Host "  VehicleBookingHistory  : 19 records" -ForegroundColor Gray
Write-Host "  VehicleBookingComment  : 11 records" -ForegroundColor Gray
Write-Host ""

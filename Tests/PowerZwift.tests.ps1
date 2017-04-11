# Tests for PowerZwift
ipmo Pester
ipmo PowerZwift

Describe "Event Related Functions" {
    $events = Get-ZwiftEvent
    Context "Without an event ID" {
        It "Should return an object array" {
            
            $events.GetType() | select -expand Name | SHould Be "Object[]"
        }
    }

    Context "with an event ID" {
        $eventID = $events | select -first 1 | select -expand id

        $eventDetails = Get-ZwiftEvent -eventID $eventID

        It "Should contain a name Property" {
            $eventdetails.name | should Not Be $null
        }
    }

}

Describe "Set-ZwiftCourse" {
    Context "Setting to default" {

    }
}

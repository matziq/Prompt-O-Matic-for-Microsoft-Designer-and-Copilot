Add-Type -AssemblyName System.Windows.Forms

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Prompt-O-Matic for Microsoft Designer'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.AutoSize = $true

# Create filter textboxes for each dropdown
$subjectFilterTextbox = New-Object System.Windows.Forms.TextBox
$subjectFilterTextbox.Location = New-Object System.Drawing.Point(200, 30)
$styleFilterTextbox = New-Object System.Windows.Forms.TextBox
$styleFilterTextbox.Location = New-Object System.Drawing.Point(200, 70)
$mediaFilterTextbox = New-Object System.Windows.Forms.TextBox
$mediaFilterTextbox.Location = New-Object System.Drawing.Point(200, 110)
$actionFilterTextbox = New-Object System.Windows.Forms.TextBox
$actionFilterTextbox.Location = New-Object System.Drawing.Point(200, 150)
$extraFilterTextbox = New-Object System.Windows.Forms.TextBox
$extraFilterTextbox.Location = New-Object System.Drawing.Point(200, 190)
$colorFilterTextbox = New-Object System.Windows.Forms.TextBox
$colorFilterTextbox.Location = New-Object System.Drawing.Point(200, 230)

# Create dropdown for Subject
$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Text = 'Subject (*Required):'
$subjectLabel.AutoSize = $true
$subjectLabel.Location = New-Object System.Drawing.Point(10, 10)
$subjectDropdown = New-Object System.Windows.Forms.ComboBox
$subjectDropdown.Location = New-Object System.Drawing.Point(10, 30)
$subjectDropdown.Items.AddRange(@('Abyss', 'Action', 'Actor', 'Adventure', 'Alien', 'Animal', 'Apartment', 'Appliance', 'Artist', 'Athlete', 'Baby', 'Bar', 'Beach', 'Bear', 'Bedroom', 'Biopic', 'Brotherhood', 'Building', 'Businessperson', 'Castle', 'Cat', 'Cave', 'Chef', 'Church', 'City', 'Club', 'Conflict', 'Continent', 'Country', 'Cowboy', 'Crime', 'Criminal', 'Dancer', 'Desert', 'Detective', 'Devil', 'Dining', 'Documentary', 'Drama', 'Drink', 'Earth', 'Equality', 'Family', 'Fight', 'Flower', 'Food', 'Forest', 'Fraternity', 'Fruit', 'Furniture', 'Galaxy', 'Girl', 'Grave', 'Heaven', 'Hell', 'Hero', 'Hospital', 'House', 'Island', 'Kitchen', 'Lake', 'Limbo', 'Man', 'Moon', 'Mosque', 'Mountain', 'Museum', 'Musician', 'Ninja', 'Nurse', 'Ocean', 'Office', 'Palace', 'Paradise', 'Park', 'Pirate', 'Planet', 'Plant', 'Politician', 'Purgatory', 'Restaurant', 'River', 'Robot', 'Room', 'Ruin', 'Rural', 'Sasquatch', 'School', 'Scientist', 'Shrine', 'Sisterhood', 'Soldier', 'Sorcerer', 'Spy', 'Star', 'Store', 'Student', 'Synagogue', 'Teacher', 'Temple', 'Theater', 'Thriller', 'Tomb', 'Tool', 'Town', 'Tree', 'Universe', 'Vampire', 'Vehicle', 'Village', 'Villain', 'Void', 'Werewolf', 'Western', 'Witch', 'Wizard', 'Woman', 'World', 'Writer'))

# Create dropdown for Action
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = 'Action (*Required):'
$actionLabel.AutoSize = $true
$actionLabel.Location = New-Object System.Drawing.Point(10, 130)
$actionDropdown = New-Object System.Windows.Forms.ComboBox
$actionDropdown.Location = New-Object System.Drawing.Point(10, 150)
$actionDropdown.Items.AddRange(@('agreeing', 'arguing', 'bargaining', 'betraying', 'blackmailing', 'bribing', 'bullying', 'buying', 'celebrating', 'chatting', 'cheating', 'cleaning', 'climbing', 'competing', 'compromising', 'conspiring', 'cooking', 'cooperating', 'corrupting', 'dancing', 'debating', 'decorating', 'defending', 'destroying', 'discussing', 'drinking', 'driving', 'eating', 'editing', 'escaping', 'exercising', 'experimenting', 'explaining', 'failing', 'fighting', 'filming', 'fishing', 'flying', 'gardening', 'harassing', 'hiding', 'hunting', 'intimidating', 'investing', 'jumping', 'learning', 'listening', 'logging', 'losing', 'lying', 'manufacturing', 'meditating', 'meeting', 'mining', 'mourning', 'murdering', 'negotiating', 'painting', 'photographing', 'planning', 'playing', 'plotting', 'praying', 'processing', 'protecting', 'reading', 'relaxing', 'repairing', 'researching', 'riding', 'rolling', 'running', 'saving', 'scheming', 'sculpting', 'selling', 'shopping', 'singing', 'sitting', 'sleeping', 'smoking', 'spending', 'standing', 'stealing', 'studying', 'succeeding', 'swimming', 'talking', 'teaching', 'thinking', 'training', 'trading', 'trapping', 'traveling', 'visiting', 'walking', 'wasting', 'watching', 'winning', 'working', 'writing'))

# Create dropdown for Style
$styleLabel = New-Object System.Windows.Forms.Label
$styleLabel.Text = 'Style (*Required):'
$styleLabel.AutoSize = $true
$styleLabel.Location = New-Object System.Drawing.Point(10, 50)
$styleDropdown = New-Object System.Windows.Forms.ComboBox
$styleDropdown.Location = New-Object System.Drawing.Point(10, 70)
$styleDropdown.Items.AddRange(@('4k resolution','8k resolution','abstract','abstract expressionist','action','action painter','anime','cartoon','color field','comic book','computer-generated','cubist','cyberpunk','dadaist','digital','documentary','dystopian','expressionistic','extremely detailed','fantasy','far eastern','fauvist','film','fine art','futuristic','graffiti','hand-drawn','hard edge','horror','hyper-realistic','hyperrealist','impressionistic','industrial','low poly','manga','mechanical','minimalist','modern','natural','noir','op art','organic','period','photo-realistic','photorealist','pointillist','pop art','postapocalyptic','realist','realistic','rural','sci-fi','soft edge','steampunk','street art','studio extra','surreal','surreal fantastic realist','surreal realist','surrealist','superrealist','thriller','traditional','urban','utopian','visionary realist','western'))

# Create dropdown for Media
$mediaLabel = New-Object System.Windows.Forms.Label
$mediaLabel.Text = 'Media (Optional):'
$mediaLabel.AutoSize = $true
$mediaLabel.Location = New-Object System.Drawing.Point(10, 90)
$mediaDropdown = New-Object System.Windows.Forms.ComboBox
$mediaDropdown.Location = New-Object System.Drawing.Point(10, 110)
$mediaDropdown.Items.AddRange(@('painting', 'sculpture', 'photography', 'printmaking', 'drawing', 'digital art', 'film', 'performance art', 'ceramics', 'fiber art', 'glass art', 'jewelry', 'metalworking', 'woodworking', 'mixed media', 'installation art', 'land art', 'conceptual art', 'video art', 'sound art', 'textile arts', 'mosaic', 'calligraphy', 'graffiti', 'comic art', 'sand art', 'ice sculpture', 'origami', 'shadow art', '3d printing art', 'bio art', 'eco art', 'food art', 'neon art', 'paper art', 'street art', 'tape art', 'body art', 'light art', 'miniature art', 'artistic bookbinding', 'pyrography', 'art toys', 'balloon art', 'latte art', 'nail art', 'makeup art', 'floral design', 'sandcastle building', 'snow sculpting', 'topiary', 'toy design', 'virtual art', 'augmented reality art', 'virtual reality art', 'algorithmic art', 'ascii art', 'biofeedback art', 'fractal art', 'generative art', 'glitch art', 'holography', 'interactive art', 'kinetic art', 'laser art', 'light painting', 'mathematical art', 'optical art', 'robotic art', 'software art', 'space art', 'synesthetic art', 'vaporwave art', 'virtual world art', 'visionary art', 'wearable art'))

# Create dropdown for Extra
$extraLabel = New-Object System.Windows.Forms.Label
$extraLabel.Text = 'Extra (Optional):'
$extraLabel.AutoSize = $true
$extraLabel.Location = New-Object System.Drawing.Point(10, 170)
$extraDropdown = New-Object System.Windows.Forms.ComboBox
$extraDropdown.Location = New-Object System.Drawing.Point(10, 190)
# Add more extra words here
$extraDropdown.Items.AddRange(@('abstract', 'art deco', 'art nouveau', 'baroque', 'classicism', 'conceptual art', 'constructivism', 'cubism', 'dadaism', 'expressionism', 'fauvism', 'futurism', 'impressionism', 'minimalism', 'modernism', 'neo-classicism', 'neo-expressionism', 'op art', 'pop art', 'post-impressionism', 'realism', 'renaissance', 'rococo', 'romanticism', 'surrealism', 'symbolism', 'vorticism'))

# Create dropdown for Color
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = 'Color Palette (Optional):'
$colorLabel.AutoSize = $true
$colorLabel.Location = New-Object System.Drawing.Point(10, 210)
$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Location = New-Object System.Drawing.Point(10, 230)
# Add more extra words here
$colorDropdown.Items.AddRange(@('autumn color palette', 'winter color palette','spring color palette','summer color palette','black and white','monochrome','sepia','pastel','neon','primary colors','secondary colors','tertiary colors','complementary colors','analogous colors','triadic colors','split-complementary colors','tetradic colors','warm colors','cool colors','neutral colors','earth tones','jewel tones','bright colors','dark colors','light colors','muted colors','saturated colors','desaturated colors','vibrant colors','dull colors','light color palette','dark color palette','bright color palette','muted color palette','saturated color palette','desaturated','color palette','vibrant color palette','dull color palette','light color scheme','dark color scheme','bright color scheme','muted color scheme','saturated color scheme','desaturated color scheme','vibrant color scheme','dull color scheme','light color combination','dark color combination','bright color combination','muted color combination','saturated color combination','desaturated color combination','vibrant color combination','dull color combination'))

# Set the 'DropDownStyle' property to 'DropDownList' for the first three dropdowns to make them mandatory
$subjectDropdown.DropDownStyle = 'DropDownList'
$styleDropdown.DropDownStyle = 'DropDownList'
$mediaDropdown.DropDownStyle = 'DropDownList'

# Add a default 'Select' item at the beginning of the list for the first three dropdowns
$subjectDropdown.Items.Insert(0, '*')
$styleDropdown.Items.Insert(0, '*')
$mediaDropdown.Items.Insert(0, '*')

# Set the 'DropDownStyle' property to 'DropDown' for the last three dropdowns to make them optional
$actionDropdown.DropDownStyle = 'DropDown'
$extraDropdown.DropDownStyle = 'DropDown'
$colorDropdown.DropDownStyle = 'DropDown'

# Enable auto-complete mode for dropdowns
$subjectDropdown.DropDownStyle = 'DropDown'
$subjectDropdown.AutoCompleteMode = 'SuggestAppend'
$subjectDropdown.AutoCompleteSource = 'ListItems'
$styleDropdown.DropDownStyle = 'DropDown'
$styleDropdown.AutoCompleteMode = 'SuggestAppend'
$styleDropdown.AutoCompleteSource = 'ListItems'
$mediaDropdown.DropDownStyle = 'DropDown'
$mediaDropdown.AutoCompleteMode = 'SuggestAppend'
$mediaDropdown.AutoCompleteSource = 'ListItems'
$actionDropdown.DropDownStyle = 'DropDown'
$actionDropdown.AutoCompleteMode = 'SuggestAppend'
$actionDropdown.AutoCompleteSource = 'ListItems'
$extraDropdown.DropDownStyle = 'DropDown'
$extraDropdown.AutoCompleteMode = 'SuggestAppend'
$extraDropdown.AutoCompleteSource = 'ListItems'
$colorDropdown.DropDownStyle = 'DropDown'
$colorDropdown.AutoCompleteMode = 'SuggestAppend'
$colorDropdown.AutoCompleteSource = 'ListItems'

# Create button to copy to clipboard
$button = New-Object System.Windows.Forms.Button
$button.Size = New-Object System.Drawing.Size(75, 50)
$button.Text = 'Copy to Clipboard'
$button.Location = New-Object System.Drawing.Point(20, 250)

$button.Add_Click({
    $selectedWords = @()
    if ($subjectDropdown.SelectedItem) { $selectedWords += $subjectDropdown.SelectedItem }
    if ($styleDropdown.SelectedItem) { $selectedWords += $styleDropdown.SelectedItem }
    if ($mediaDropdown.SelectedItem) { $selectedWords += $mediaDropdown.SelectedItem }
    if ($actionDropdown.SelectedItem) { $selectedWords += $actionDropdown.SelectedItem }
    if ($extraDropdown.SelectedItem) { $selectedWords += $extraDropdown.SelectedItem }
    if ($colorDropdown.SelectedItem) { $selectedWords += $colorDropdown.SelectedItem }
    $clipboardText = [string]::Join(', ', $selectedWords)
    $clipboardText = $clipboardText.Substring(0,1).ToUpper()+$clipboardText.Substring(1) + "."
    Set-Clipboard -Value $clipboardText
    $copiedTextbox.Text = $clipboardText
})

# Create label and textbox to display copied text
$copiedLabel = New-Object System.Windows.Forms.Label
$copiedLabel.Text = 'Copied Text:'
$copiedLabel.AutoSize = $true
$copiedLabel.Location = New-Object System.Drawing.Point(10, 300)
$copiedTextbox = New-Object System.Windows.Forms.TextBox
$copiedTextbox.Location = New-Object System.Drawing.Point(10, 340)
$copiedTextbox.ReadOnly = $true

# Create "Random" button
$randomButton = New-Object System.Windows.Forms.Button
$randomButton.Text = 'Random Prompt'
$randomButton.Location = New-Object System.Drawing.Point(150, 250)
$randomButton.Size = New-Object System.Drawing.Size(75, 50)
$randomButton.Add_Click({
    $selectedWords = @()
    $selectedWords += $subjectDropdown.Items | Get-Random
    $selectedWords += $styleDropdown.Items | Get-Random
    $selectedWords += $mediaDropdown.Items | Get-Random
    $selectedWords += $actionDropdown.Items | Get-Random
    $selectedWords += $extraDropdown.Items | Get-Random
    $selectedWords += $colorDropdown.Items | Get-Random
    $clipboardText = [string]::Join(', ', $selectedWords)
    if ($clipboardText) {
        $clipboardText = $clipboardText.Substring(0,1).ToUpper()+$clipboardText.Substring(1) + "."
        Set-Clipboard -Value $clipboardText
        $copiedTextbox.Text = $clipboardText
    }
})

# Add a button to open the website with the generated text
$webButton = New-Object System.Windows.Forms.Button
$webButton.Text = 'Open in Microsoft Designer'
$webButton.Location = New-Object System.Drawing.Point(290, 250)
$webButton.Size = New-Object System.Drawing.Size(75, 50)
$webButton.Add_Click({
    # Get the text from the clipboard
    $clipboardText = Get-Clipboard
    if ($clipboardText) {
        # Format the text for the URL
        $urlText = Format-ForURL($clipboardText)
        # Open the website with the generated text
        Start-Process "$baseURL$urlText"
    }
})

# Set the copied text to the joined selected words
if ($selectedWords -ne $null) {$copiedTextbox.Text = $selectedWords -join ' '} else {$copiedTextbox.Text = ""}

# Define the base URL for the website
$baseURL = "https://designer.microsoft.com/image-creator?p="
# Create a function to format the text for the URL
function Format-ForURL ($text) {
    $text = $text.Replace(', ', '+').Replace(' ', '+')
    return $text
}

# Add a button to open the website with the generated text
$webButton = New-Object System.Windows.Forms.Button
$webButton.Text = 'Open in Microsoft Designer'
$webButton.Location = New-Object System.Drawing.Point(290, 250)
$webButton.Size = New-Object System.Drawing.Size(75, 50)
$webButton.Add_Click({

    # Capture the currently chosen items from the dropdowns
    $selectedWords = @()
    if ($subjectDropdown.SelectedItem) { $selectedWords += $subjectDropdown.SelectedItem }
    if ($styleDropdown.SelectedItem) { $selectedWords += $styleDropdown.SelectedItem }
    if ($mediaDropdown.SelectedItem) { $selectedWords += $mediaDropdown.SelectedItem }
    if ($actionDropdown.SelectedItem) { $selectedWords += $actionDropdown.SelectedItem }
    if ($extraDropdown.SelectedItem) { $selectedWords += $extraDropdown.SelectedItem }
    if ($colorDropdown.SelectedItem) { $selectedWords += $colorDropdown.SelectedItem }
    $clipboardText = [string]::Join(', ', $selectedWords)
    if ($clipboardText) {
        $clipboardText = $clipboardText.Substring(0,1).ToUpper()+$clipboardText.Substring(1) + "."
        Set-Clipboard -Value $clipboardText
        $copiedTextbox.Text = $clipboardText
    }

    # Open the website with the generated text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    $url = $baseURL + $formattedText
    Start-Process $url
})


# Add the selected items from the Dropdowns to the selectedWords array
if ($colorDropdown.Text) { $selectedWords += $colorDropdown.Text }
if ($extraDropdown.Text) { $selectedWords += $extraDropdown.Text }
if ($actionDropdown.Text) { $selectedWords += $actionDropdown.Text }
if ($mediaDropdown.Text) { $selectedWords += $mediaDropdown.Text }
if ($styleDropdown.Text) { $selectedWords += $styleDropdown.Text }
if ($subjectDropdown.Text) { $selectedWords += $subjectDropdown.Text }

# Add the web button to the form controls
$form.Controls.Add($webButton)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Add the random button to the form controls
$form.Controls.Add($randomButton)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Add controls to form
$form.Controls.AddRange(@($subjectLabel, $subjectDropdown, $actionLabel, $actionDropdown, $styleLabel, $styleDropdown, $mediaLabel, $mediaDropdown, $extraLabel, $extraDropdown, $colorLabel, $colorDropdown, $button, $copiedLabel, $copiedTextbox))
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Set the size of the copiedTextbox to be large enough to display the entire set of words
$copiedTextbox.Size = New-Object System.Drawing.Size(500, 100)

# Show form
$form.ShowDialog()
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# SIG # Begin signature block
# MIIJCQYJKoZIhvcNAQcCoIII+jCCCPYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUznP/xRirwiC4RwehpJl6bv3A
# jWSgggZqMIIGZjCCBU6gAwIBAgITJAAAAUPgt24PVLctvwAAAAABQzANBgkqhkiG
# 9w0BAQsFADBRMRMwEQYKCZImiZPyLGQBGRYDb3JnMRwwGgYKCZImiZPyLGQBGRYM
# bWVtY29uZmlnbWdyMRwwGgYDVQQDExNtZW1jb25maWdtZ3ItQ00xLUNBMB4XDTI0
# MDQxNzE0MzI1OVoXDTI2MDQxNzE0MzI1OVowEzERMA8GA1UEAxMIbGFiYWRtaW4w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDv2ENL7267CRUs4KG//Mnf
# UQj9JKkVF0SZOLZFWRxkc7md06Ie3sPujBJ8WQDRZGex6DvmucHPDGKC8yxQRnbc
# /i1KH9k2UX0U3baag0CWSGoBgZyBuTMY5/yIFiQEJsEue6W51kN9LyDgvvbMNLi6
# n3Typ5e4VgfqlQay6HtJSeZet5s2BQMyOt3LjBHs/SSpTCMzk2/DrR5lf4gIfNb/
# xI6loNzlhTZ1WCurMES8UL0jRyvsqYsEJ+1fCD8tXXVjxWxOdUXAUjpPtkJrwhBC
# PUIyXIAF/PTvYUVu1KwYu7MKuH0Jav0lTH9pFWndDRC7/3bsXHREf53bgtmFXgep
# AgMBAAGjggNzMIIDbzA8BgkrBgEEAYI3FQcELzAtBiUrBgEEAYI3FQjPpDWH0vEw
# hPWJNoOEnwCGjaALXIHCphqGrf1fAgFsAgEAMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MA4GA1UdDwEB/wQEAwIHgDAbBgkrBgEEAYI3FQoEDjAMMAoGCCsGAQUFBwMDMB0G
# A1UdDgQWBBTEh+3Hi6JSfKfq9NrAKyDCC9JyLzAfBgNVHSMEGDAWgBSEq85L/ec2
# GZ2gu1TNr8SgdoHScDCCAVkGA1UdHwSCAVAwggFMMIIBSKCCAUSgggFAhoG7bGRh
# cDovLy9DTj1tZW1jb25maWdtZ3ItQ00xLUNBLENOPUNNMSxDTj1DRFAsQ049UHVi
# bGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlv
# bixEQz1tZW1jb25maWdtZ3IsREM9b3JnP2NlcnRpZmljYXRlUmV2b2NhdGlvbkxp
# c3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2ludIY+aHR0cDov
# L0NNMS5tZW1jb25maWdtZ3Iub3JnL0NlcnRFbnJvbGwvbWVtY29uZmlnbWdyLUNN
# MS1DQS5jcmyGQGZpbGU6Ly8vL0NNMS5tZW1jb25maWdtZ3Iub3JnL0NlcnRFbnJv
# bGwvbWVtY29uZmlnbWdyLUNNMS1DQS5jcmwwgcoGCCsGAQUFBwEBBIG9MIG6MIG3
# BggrBgEFBQcwAoaBqmxkYXA6Ly8vQ049bWVtY29uZmlnbWdyLUNNMS1DQSxDTj1B
# SUEsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29u
# ZmlndXJhdGlvbixEQz1tZW1jb25maWdtZ3IsREM9b3JnP2NBQ2VydGlmaWNhdGU/
# YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MDQGA1UdEQQt
# MCugKQYKKwYBBAGCNxQCA6AbDBlsYWJhZG1pbkBtZW1jb25maWdtZ3Iub3JnME0G
# CSsGAQQBgjcZAgRAMD6gPAYKKwYBBAGCNxkCAaAuBCxTLTEtNS0yMS0yNzU3MzIw
# NDAtOTUyNDUzNzU0LTI4ODk3NTA5NTAtMTEwMzANBgkqhkiG9w0BAQsFAAOCAQEA
# eoRuRgzEzrTSPnayr1mefYvXU8NRQcRUzZs4SYbs7cINqKK17EY5trDLrAR9FsNP
# Ya4ZcWa6nyaciK3iTZjeeqcHujmqA9viEMO/bpdak2yNfu60fwuH7OhHXcCYZf1n
# kCE7m2QuM5WJ2WWhNwx+lHKbvYs5z0DHweOUAe8kF5aVYjPXFbQEi7VFPhXJK7KV
# D/cFltTQgEwHowMysrnEYr3CHIf2HP4K2riL+5Wh5H0LX7af7HdRmo8jQnEHkYbH
# Zn+gilKGk6TXsiUVv5eqPPIGtTi6yA9m9dNbOslf9iK0oeISv2JJ0wGXSmOte+X7
# J4tH5/slqyFpky+23p0o0TGCAgkwggIFAgEBMGgwUTETMBEGCgmSJomT8ixkARkW
# A29yZzEcMBoGCgmSJomT8ixkARkWDG1lbWNvbmZpZ21ncjEcMBoGA1UEAxMTbWVt
# Y29uZmlnbWdyLUNNMS1DQQITJAAAAUPgt24PVLctvwAAAAABQzAJBgUrDgMCGgUA
# oHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYB
# BAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0B
# CQQxFgQU1vuZgPpmLub1tKx/xSHhMybmZ18wDQYJKoZIhvcNAQEBBQAEggEAQ8Ed
# Wg+d2lOexrESyUxETkiaJxEMFd5zesKYsismKmGON2UMpDyJZRISb7qj7/JHBJpv
# Dr2vK0JpD9RdKp8b7od5enMTV+OnY+jnyx8pnFAsSWCU03qWIdNG4NTQX/RH6oMF
# LMiyk+8EyOL4DJjL6SQIX6YVQlfgwKqAqKYVYwxssQFUV8mKocnLfyY/164A1ail
# 3YtU4P3y+6+XLDKI2ngbzYY4/uVojHNfg9IC1FLMnkJbBK59M7h8EyxRGxGBN08k
# 7WcNQSj27mC6co8EDvmgC+JJXuI31xNtgzlWa5ShoS2lRGXwmkrcY+IHCl6AFfPE
# py8aeKIUtAppd+Oimw==
# SIG # End signature block

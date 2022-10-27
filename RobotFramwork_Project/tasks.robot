*** Settings ***
Documentation       Etsi sopivia kokeita ja listaa ne Excel tiedostoon

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.JavaAccessBridge
Library             RPA.Desktop.Windows
Library             RPA.Windows
Library             Collections


*** Variables ***
${Tuomari}      /html/body/div/div[1]/div[2]/div/div/div/div/table/tbody/tr[3]/td/div[10]/div/div
@{Tiedosto}     Ilmoittaudu_kokeisiin.xlsx


*** Tasks ***
Päätehtävä 1: Avaa verkkosivu ja täytä lomake
    Open Available Browser
    Toggle Drop Down    //*[@id="lajiSel"]    //*[@id="lajiSel"]/option[15]
    RPA.Browser.Selenium.Click Element    //*[@id="koeluokka53"]
    Click Button    //*[@id="listaaBtn"]

Päätehtävä 2: Kerää kriteereihin sopivat rally-toko kokeet ja laita excel tiedostoon
    Menu Select    //*[@id="hakutulos"]/table/tbody/tr[2]/td[2]/a
    Select
    ...    //*[@id="hakutulos"]/table/tbody/tr[2]/td[3]
    ...    document.querySelector("#hakutulos > table > tbody > tr:nth-child(2) > td:nth-child(3)")
    Open Available Browser
    ...    https://www.google.com/maps/dir/Maauunintie+27,+Vantaa//@60.3440623,24.9965856,12z/data=!4m9!4m8!1m5!1m1!1s0x4692074661f2d40b:0xfd2e736c79382a6!2m2!1d25.0666248!2d60.3440807!1m0!3e0
    Input Text
    ...    //*[@id="sb_ifc52"]/input
    ...    text=document.querySelector("#hakutulos > table > tbody > tr:nth-child(2) > td:nth-child(3)")
    Click Button    hakuBt
    Select
    ...    //*[@id="section-directions-trip-0"]/div[1]/div[1]/div[1]/div[2]/div
    ...    document.querySelector("#section-directions-trip-0 > div.MespJc > div:nth-child(1) > div.XdKEzd > div.ivN21e.tUEI8e.fontBodyMedium > div")
    @{Välimatkalista}=    Create List
    IF    ${document.querySelector("#section-directions-trip-0 > div.MespJc > div:nth-child(1) > div.XdKEzd > div.ivN21e.tUEI8e.fontBodyMedium > div")} > ${100}
        Append To List    ${Välimatkalista}
    ELSE
        Remove From List    ${Välimatkalista}    ${i}
    END
    IF    ${Tuomari} == ${'Riikka Timonen'}
        Append To List    ${Välimatkalista}
    ELSE
        Remove From List    ${Välimatkalista}    ${i}
    END
    Log To Console    ${item}

Päätehtävä 3: Lajittele tulokset

[Teardown] Sulje verkkosivu ja mahdolliset excel-tiedostot

Ongelmatilanne: Ilmoittaa ongelmasta ja pyytää ihmisen apua


*** Keywords ***
Päätehtävä 1: Avaa verkkosivu ja täytä lomake
    Open Available Browser    https://www.virkku.net/index.cfm?template=search.cfm&type=kokeet
    Toggle Drop Down    //*[@id="lajiSel"]    //*[@id="lajiSel"]/option[15]
    RPA.Browser.Selenium.Click Element    //*[@id="koeluokka53"]
    Click Button    //*[@id="listaaBtn"]
    #Toggle Drop Down    //*[@id="lajiSel"]    //*[@id="lajiSel"]/option[16]
    #Jätettiin toko-lajivaihtoehto pois

Päätehtävä 2: Kerää kriteereihin sopivat rally-toko kokeet ja laita excel tiedostoon
    Menu Select    //*[@id="hakutulos"]/table/tbody/tr[2]/td[2]/a
    Select
    ...    //*[@id="hakutulos"]/table/tbody/tr[2]/td[3]
    ...    document.querySelector("#hakutulos > table > tbody > tr:nth-child(2) > td:nth-child(3)")
    Open Available Browser
    ...    https://www.google.com/maps/dir/Maauunintie+27,+Vantaa//@60.3440623,24.9965856,12z/data=!4m9!4m8!1m5!1m1!1s0x4692074661f2d40b:0xfd2e736c79382a6!2m2!1d25.0666248!2d60.3440807!1m0!3e0
    Input Text
    ...    //*[@id="sb_ifc52"]/input
    ...    text=document.querySelector("#hakutulos > table > tbody > tr:nth-child(2) > td:nth-child(3)")
    Click Button    hakuBt
    Select
    ...    //*[@id="section-directions-trip-0"]/div[1]/div[1]/div[1]/div[2]/div
    ...    document.querySelector("#section-directions-trip-0 > div.MespJc > div:nth-child(1) > div.XdKEzd > div.ivN21e.tUEI8e.fontBodyMedium > div")
    @{Välimatkalista}=    Create List
    IF    ${document.querySelector("#section-directions-trip-0 > div.MespJc > div:nth-child(1) > div.XdKEzd > div.ivN21e.tUEI8e.fontBodyMedium > div")} > ${100}
        Append To List    ${Välimatkalista}
    ELSE
        Remove From List    ${Välimatkalista}    ${i}
    END
    IF    ${Tuomari} == ${'Riikka Timonen'}
        Append To List    ${Välimatkalista}
    ELSE
        Remove From List    ${Välimatkalista}    ${i}
    END
    Log To Console    ${item}

    # Vaihtoehtoinen toteutusidea
    #Input Text    ${paikkakunta}    (id tähän)
    #Skip if matka > 100 km
    #FOR ${matka} IN (//*[@id="section-directions-trip-0"]/div[1]/div[1]/div[1]/div[2]/div)
    #IF $maaranpaa > 100 km CONTINUE
    #END

    #Skip if tuomari Riikka Timonen jne.
    #FOR ${koe} IN @{"C:\Users\Omistaja\OneDrive\Tiedostot\Ohjelmistorobotiikka (RPA)\Ilmoittaudu kokeisiin.xlsx"}(?)
    #IF $tuomari == 'Riikka Timonen '    CONTINUE
    #IF $tuomari == 'Susanna Berghäll'    CONTINUE
    #IF $tuomari == 'Iiris Harju'    CONTINUE
    #IF $tuomari == 'Anna Pekkala'    CONTINUE
    #END

Päätehtävä 3: Lajittele tulokset

Avaa Excel tiedosto Ilmoittaudu kokeisiin
    Create Workbook    Ilmoittaudu_kokeisiin.xlsx    overwrite=True
    Open Workbook    Ilmoittaudu_kokeisiin.xlsx    overwrite=True

Listaa koe ja ilmoittautumispäivämäärät
    #${koepäivämäärä} ${ilmoittautumispäivämäärä}
    FOR ${koe} IN    @{Tiedosto}(Ilmoittaudu_kokeisiin.xlsx)
    Set Worksheet Value    first available    first available    value
    Save Workbook

Sulje verkkosivu ja mahdolliset excel-tiedostot
    Close Browser
    Close Workbook

Ongelmatilanne: Ilmoittaa ongelmasta ja pyytää ihmisen apua
    Run Keyword And Expect Error    exit status 5    Log To Console    Ongelma robotin koodissa.
    Run Keyword And Expect Error    ValueError: *    Log To Console    Ongelma arvojen saamisessa.
    Run Keyword And Continue On Failure    Log to console    Apua, tuli ongelma, enkä osaa edetä.

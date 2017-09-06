function viewRecords(id) {
    // set up objects 
    var r1 = document.getElementById('record_1');
    var r2 = document.getElementById('record_2');
    var h1 = document.getElementById('compare');
    var headers = document.getElementById('header').getElementsByTagName('td');
    var r1Data = document.getElementById('r' + id).getElementsByTagName('td');
    var newel = document.createElement('td');
    var arLen = (document.getElementById('r' + id).getElementsByTagName('td').length - 1) / 2
    var spacer = document.createElement('td');

    // create arrays
    var arr_1 = [arLen * 2]
    var arr_2 = [arLen * 2]

    // clean up the comparision area
    while (h1.firstChild) {
        h1.removeChild(h1.firstChild);
    }
    while (r1.firstChild) {
        r1.removeChild(r1.firstChild);
    }
    while (r2.firstChild) {
        r2.removeChild(r2.firstChild);
    }

    // populate arrays
    for (ele = 0; ele < arLen; ele++) {
        arr_1[ele] = document.getElementById('r' + id).getElementsByTagName('td')[ele + 1].innerHTML;
    }
    for (ele = 0; ele < arLen; ele++) {
        arr_2[ele] = document.getElementById('r' + id).getElementsByTagName('td')[(ele + arLen + 1)].innerHTML;
    }

    //h1.appendChild(spacer).setAttribute('style','border:0');
    
    // add the headers
    for (el = 1; el < arLen; el++) {

        window['hr' + el] = document.createElement('th')
        window['hr' + el].innerHTML = headers[el].innerHTML
        h1.appendChild(window['hr' + el]).innerHTML = headers[el].innerHTML
    }

    // add the td's to the compare area
    for (ele = 0; ele < arLen; ele++) {
        a = ele
        window['newel' + ele + '_1'] = document.createElement('td');
        window['newel' + ele + '_2'] = document.createElement('td');
        if (arr_1[ele] == arr_2[ele]) {
            r1.appendChild(window['newel' + ele + '_1']).innerHTML = arr_1[ele]
            r1.appendChild(window['newel' + ele + '_1']).setAttribute('class', 'tr-match')
            r1.appendChild(window['newel' + ele + '_1']).setAttribute('style', 'border:0')
            r2.appendChild(window['newel' + ele + '_2']).innerHTML = arr_2[ele]
            r2.appendChild(window['newel' + ele + '_2']).setAttribute('class', 'tr-match')
            r2.appendChild(window['newel' + ele + '_2']).setAttribute('style', 'border:0')
        }
        else {
            r1.appendChild(window['newel' + ele + '_1']).innerHTML = arr_1[ele]
            r1.appendChild(window['newel' + ele + '_1']).setAttribute('class', 'tr-nomatch')
            r1.appendChild(window['newel' + ele + '_1']).setAttribute('style', 'border:0')
            r2.appendChild(window['newel' + ele + '_2']).innerHTML = arr_2[ele]
            r2.appendChild(window['newel' + ele + '_2']).setAttribute('class', 'tr-nomatch')
            r2.appendChild(window['newel' + ele + '_2']).setAttribute('style', 'border:0')
        }
    }
	
    //reload();
    //console.log(arr_1.toString())
    //console.log(arr_2.toString())
}

function reload() {
    var container = document.getElementById('compare');
    var content = container.innerHTML;
    container.innerHTML = '';
    container.innerHTML = content;
}

function getFile(input) {
    let file = input.files[0]
    if (file.type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
        showNotification({message: 'Неверный формат файла!', style: 'red'})
        return
    }
    document.getElementById('openBtn').disabled = true
    document.getElementById('button_block').style.display = 'block'

    document.getElementById('fileName').innerHTML = file.name
}


function sendFile(){
    document.getElementById('sendBtn').disabled = true
    let file = document.getElementById('file').files[0]
    let formData = new FormData()
    formData.append('file', file)

    fetch('', {
        method: 'POST',
        body: formData
    }).then(response => {
        if (!response.ok) {
            showNotification({message: 'Что-то пошло не так...', style: 'failed'})
        } else {
            showNotification({message: 'OK', style: 'success'})
        }
        return response.blob()
    })
    .then(blob => {
        document.getElementById('button_block').style.display = 'none'

        let downloadUrl = window.URL.createObjectURL(blob)
        let a = document.createElement('a')
        a.href = downloadUrl
        a.download = 'Отчет.xlsx'
        document.body.appendChild(a)
        a.click()

        setTimeout(() => {
            window.location = '/'
        }, 2000)
    })

}



// всплывающее уведомление
function showNotification({message, style}) {
    let notification = document.createElement('div')

    if ((style == 'green') || (style == 'success')) {
      notification.className = 'notifGreen'
    }
    if ((style == 'red') || (style == 'failed')) {
      notification.className = 'notifRed'
    }
    if (style == 'yellow') {
      notification.className = 'notifYellow'
    }
    notification.innerHTML = message

    document.body.append(notification)

    setTimeout(() => notification.remove(), 3000)
}
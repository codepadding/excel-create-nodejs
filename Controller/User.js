const User = require('../Models/User')
const excelJS = require('exceljs')

const exportUser = async (req, res) => {
    const workbook = new excelJS.Workbook()
    const worksheet = workbook.addWorksheet('My Users')
    const path = './files'

    worksheet.columns = [
        // { header: 'S no.', key: 's_no', width: 10 },
        { header: 'First Name', key: 'fname', width: 10 },
        { header: 'Last Name', key: 'lname', width: 10 },
        { header: 'Email Id', key: 'email', width: 10 },
        { header: 'Gender', key: 'gender', width: 10 },
        { header: 'Type User', key: 'type', width: 10 },
    ]

    worksheet.getCell('F1').dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"One,Two,Three,Four"'],
    }

    let counter = 1

    User.forEach(user => {
        // user.s_no = counter
        worksheet.addRow(user)
        counter++
    })

    // worksheet.getRow(1).eachCell(cell => {
    //     cell.font = { bold: true }
    // })
    worksheet
        .getColumn('E')
        .eachCell({ includeEmpty: true }, function (cell, rowNumber) {
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: User,
            }
        })

    try {
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        res.setHeader('Content-Disposition', `attachment; filename=users.xlsx`)

        return workbook.xlsx.write(res).then(() => {
            res.status(200)
        })

        // await workbook.xlsx.writeFile(`${path}/users.xlsx`).then(() => {
        //   res.send({
        //     status: "success",
        //     message: "file successfully downloaded",
        //     path: `${path}/users.xlsx`,
        //   });
        // });
    } catch (err) {
        res.send({
            status: 'error',
            message: 'Something went wrong',
        })
    }
}

module.exports = exportUser

const xlsx = require('node-xlsx')
const fse = require('fs-extra')
const inquirer = require('inquirer')
const path = require('path')
const userHome = require('user-home')

// const xlsxPath = '/Users/admin/Downloads/test.xlsx'

let xlsxPath = ''

const prdFrom = 6 // prd打分开始的索引（改动后必须修改！）
let prdEnd = 0
const gap = 2 // 2列为一个需求的分数（改动后必须修改）
const extraQuestion = 2 // prd后面多余的问题（改动后必须修改！）
let prdCount = 0

let rawData
let colNameData
const costScoreData = []

const finalResult = []

async function core() {
  await invoke()
}

core()

async function invoke() {

  const askForFilePath = async () => {
    xlsxPath = path.resolve(userHome, 'Downloads', 'data.xlsx')
    const { isUseDefaultXlsxPath } = await inquirer.prompt({
      type: 'confirm',
      name: 'isUseDefaultXlsxPath',
      message: `默认使用${xlsxPath}文件，是否继续`,
      default: true
    })

    if (isUseDefaultXlsxPath) {
      console.log('将使用默认用户下载文件下的data.xlsx文件: ' + xlsxPath)
    } else {
      const { newPath } = await inquirer.prompt({
        type: 'input',
        name: 'newPath',
        message: '请输入你的xlsx文件所在路径和文件名，类似：/a/b/c.xlsx'
      })

      xlsxPath = newPath
      console.log('将使用自定义的文件路径：' + xlsxPath)
    }

    if(!fse.existsSync(xlsxPath)) {
      throw new Error('xlsx文件不存在，请检查')
    }
  }

  const fetchData = () => {
    rawData = xlsx.parse(xlsxPath)[0].data

    colNameData = rawData[0]
    prdEnd = colNameData.length - extraQuestion
    prdCount = (prdEnd - prdFrom) / gap
  }
    
  const askForOtherData = async () => {
    const { neeInputCostScores } = await inquirer.prompt({
      type: 'confirm',
      name: 'neeInputCostScores',
      message: '是否输入成本分，选择否成本分将默认都为5分',
    })

    if (neeInputCostScores) {
      for (let i = 0; i < prdCount; i++) {    
        const prdName = colNameData[i * 2 + prdFrom]

        const answer = await inquirer.prompt({
          type: 'number',
          name: 'costScore',
          message: `请输入"${prdName}"的成本分`,
        })

        constScoreData.push(answer.costScore)
      }
    } else {
      for (let i = 0; i < prdCount; i++) { 
        costScoreData.push(5)
      }
    }

    console.log('成本分数据')
    console.log(costScoreData)
  }

  const calScorePro = () => {
    const resultColNames = ['打分者']
    // 遍历需求先生成结果第一行的colNames
    for (let i = 0; i < prdCount; i++) {
      const index = i * 2 + prdFrom    
      const prdName = colNameData[index]
      resultColNames.push(prdName)
    }
    finalResult[0] = resultColNames

    // 根据人员每次打完所有需求的分
    for(let i = 1; i < rawData.length; i++) {
      const personData = rawData[i]
      const name = personData[0]

      const personResult = [name]

      for (let j = 0; j < prdCount; j++) {
        const index = j * 2 + prdFrom    
        
        const valueScore = personData[index]
        const urgencyScore = personData[index + 1]
        const costScore = costScoreData[j] || 5 

        const score = valueScore * 0.7 + urgencyScore * 0.2 + costScore * 0.1

        personResult.push(score.toFixed(2))
      }

      finalResult[i] = personResult
    }
  }

  const calScoreAvgs = () => {
    //  最后增加一行平均分
    const avgData = ['平均分']
    const persons = finalResult.length - 1

    for(let j = 0; j < prdCount; j++) {
      let sum = 0
      for(let i = 1; i < finalResult.length; i++) {
        const personData = finalResult[i]
        const score = personData[j + 1] // j + 1是因为此时的数据为，分数从1开始而不是0 ['leo', 3, 4, 5]

        sum += parseFloat(score)
      }

      avgData.push(parseFloat(sum / persons).toFixed(2))
    }
    
    finalResult.push(avgData)

    console.log('最终计算结果')
    console.log(finalResult)
  }

  const writeFile = () => {
    const buffer = xlsx.build([{ name: 'sheets', data: finalResult }])

    const rawFilePath = xlsxPath.substring(0, xlsxPath.lastIndexOf('/'))
    const rawFileName = xlsxPath.substring(xlsxPath.lastIndexOf('/') + 1).split('.')[0]
    const filePath = path.resolve(rawFilePath, rawFileName + '-result.xlsx')
    console.log('结果将写入：' + filePath)
    fse.writeFileSync(filePath, buffer)
    console.log('结果写入文件成功！')
  }

  try {
    await askForFilePath()
    fetchData()
    await askForOtherData()
    calScorePro()
    calScoreAvgs()
    writeFile()

  } catch (e) {
    console.log('发生错误!')
    console.log(e)
  } 
}


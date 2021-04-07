/*
 * Authors
 * Sarah Alshawkani
 * Mohammed Alsayed
 * Younes Alturkey
 */

Main()

function Main() {
  function start(path) {
    // check if there is a file
    if (path != null) {
      // render on existing open document
      var doc = app.documents.item(0)
      var currViewPrefs = doc.viewPreferences.properties

      // to unify the script measerment in all devices
      doc.viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.MILLIMETERS
      doc.viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.MILLIMETERS

      // Adjust the In Design Template to the required size
      doc.documentPreferences.pageWidth = '210 mm'
      doc.documentPreferences.pageHeight = '297 mm'

      var myMasterSpread = doc.masterSpreads.item(0)
      var myMarginPreferences = myMasterSpread.pages.item(0).marginPreferences

      myMarginPreferences.top = 18
      myMarginPreferences.bottom = 13
      myMarginPreferences.columnCount = 3
      myMarginPreferences.columnGutter = 67.75

      // to identfy the color for descrption for each group
      var color = doc.colors.add({
        name: 'C=0 M=0 Y=0 K=0',
        space: ColorSpace.CMYK,
        model: ColorModel.process,
        colorValue: [94, 58, 52, 37],
      })

      // iniliaze a text frame with changable size to save the data column from the excel sheet.

      var tmp_textframe = doc.pages[0].textFrames.add()
      tmp_textframe.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe1 = doc.pages[0].textFrames.add()
      tmp_textframe1.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe2 = doc.pages[0].textFrames.add()
      tmp_textframe2.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe3 = doc.pages[0].textFrames.add()
      tmp_textframe3.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe4 = doc.pages[0].textFrames.add()
      tmp_textframe4.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe5 = doc.pages[0].textFrames.add()
      tmp_textframe5.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe6 = doc.pages[0].textFrames.add()
      tmp_textframe6.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }
      var tmp_textframe7 = doc.pages[0].textFrames.add()
      tmp_textframe7.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe8 = doc.pages[0].textFrames.add()
      tmp_textframe8.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe9 = doc.pages[0].textFrames.add()
      tmp_textframe9.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      var tmp_textframe10 = doc.pages[0].textFrames.add()
      tmp_textframe10.properties = {
        textFramePreferences: {
          autoSizingReferencePoint: 1953459301,
          autoSizingType: 1752070000,
        },
      }

      //Automate Excel Sheet Cell Range Selection
      function setExcelImportPrefs(maxRange, letter) {
        app.excelImportPreferences.rangeName = letter
          .concat('1:')
          .concat(letter)
          .concat(maxRange)
        app.excelImportPreferences.tableFormatting =
          TableFormattingOptions.excelUnformattedTabbedText
      }

      // first try to check that the process of reading from excel sheet are applicable without errors
      try {
        //pick the data
        // Placing the range of the chosen data column in the frame

        // Take number of rows from user to automate cell range input A1:A[maxRange]
        maxRange = maxRange = prompt(
          'Please, enter the max number of rows (e.g. 93).',
          '',
          'Excel Sheet Max Range For Rows'
        )

        // read group code
        tmp_textframe.place(path, setExcelImportPrefs(maxRange, 'A'))

        // read item code RMC
        tmp_textframe1.place(path, setExcelImportPrefs(maxRange, 'B'))

        //read description
        tmp_textframe2.place(path, setExcelImportPrefs(maxRange, 'C'))

        // read component
        tmp_textframe3.place(path, setExcelImportPrefs(maxRange, 'D'))

        // read division
        tmp_textframe4.place(path, setExcelImportPrefs(maxRange, 'E'))

        //  read Retail price
        tmp_textframe5.place(path, setExcelImportPrefs(maxRange, 'F'))

        // read Promo price
        tmp_textframe6.place(path, setExcelImportPrefs(maxRange, 'G'))

        // read saving
        tmp_textframe7.place(path, setExcelImportPrefs(maxRange, 'H'))

        // read % discount
        tmp_textframe8.place(path, setExcelImportPrefs(maxRange, 'I'))

        // read More icons
        tmp_textframe9.place(path, setExcelImportPrefs(maxRange, 'J'))

        // read Number of Products
        tmp_textframe10.place(path, setExcelImportPrefs(maxRange, 'K'))

        // pass the data to varable, so we can process it
        var column1contentsArray = tmp_textframe.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray2 = tmp_textframe1.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray3 = tmp_textframe2.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray4 = tmp_textframe3.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray5 = tmp_textframe4.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray6 = tmp_textframe5.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray7 = tmp_textframe6.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray8 = tmp_textframe7.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray9 = tmp_textframe8.parentStory.contents.split(
          '\r'
        )
        var column1contentsArray10 = tmp_textframe9.parentStory.contents.split(
          '\r'
        )

        var column1contentsArray11 = tmp_textframe10.parentStory.contents.split(
          '\r'
        )

        //Remove frames after data extraction
        tmp_textframe.remove()
        tmp_textframe1.remove()
        tmp_textframe2.remove()
        tmp_textframe3.remove()
        tmp_textframe4.remove()
        tmp_textframe5.remove()
        tmp_textframe6.remove()
        tmp_textframe7.remove()
        tmp_textframe8.remove()
        tmp_textframe9.remove()
        tmp_textframe10.remove()

        // to read the data and pass it
        var idArr = new Array()
        var picArr = new Array()
        var desArr = new Array()
        var framesArr = new Array()
        var iconArr = new Array()
        var diviArr = new Array()
        var retailArr = new Array()
        var promoArr = new Array()
        var savingArr = new Array()
        var matchingImages = new Array()
        var precentDiscountArr = new Array()
        var MoreIconsArr = new Array()

        // temp arrayes to process and render the data
        var tempArr = new Array()
        var tempArr2 = new Array()
        var tempArr3 = new Array()
        var tempArr4 = new Array()
        var tempArr5 = new Array()
        var tempArr6 = new Array()
        var tempArr7 = new Array()
        var tempArr8 = new Array()
        var tempArr9 = new Array()
        var tempArr10 = new Array()

        //[top,left,bottom,right]
        // setting the sizes of Image frames
        var arrayOfSize = new Array()
        arrayOfSize = [
          [18, 18, 68, 74.54],
          [18, 120, 68, 181.84],
          [110, 12.7, 160, 74.54],
          [110, 120, 160, 181.84],
          [206, 12.7, 256, 74.54],
          [206, 120, 256, 181.84],
        ]

        // var arrayOfIconFrame = new Array();
        // arrayOfIconFrame = [
        //     [0, 18, 70, 50],
        //     [0, 120, 70, 130],
        //     [80, 18, 190, 50],
        //     [80, 120, 190, 130],
        //     [160, 18, 284, 50],
        //     [160, 120, 284, 130]
        // ]

        // setting the sizes of icon frames
        var arrayOfdescrptionframe = new Array()
        arrayOfdescrptionframe = [
          [87, 14, 106, 94],
          [87, 118, 106, 196],
          [187, 14, 206, 94],
          [187, 118, 206, 196],
          [270, 14, 289, 94],
          [270, 118, 289, 196],
        ]

        //read from the image folder to render it
        var fImage = Folder.selectDialog('Please, select the images folder.')
        var allFiles = fImage.getFiles()

        //read from the icon folder to render it
        var fIcone = Folder.selectDialog('Please, select the icons folder.')
        var allFilesIcone = fIcone.getFiles()

        // second try to make sure the process of manuplation the data are working
        try {
          //save the excel column in the created array
          for (var i = 1; i < column1contentsArray.length; i++) {
            idArr.push(column1contentsArray[i])
          }

          for (var i = 1; i < column1contentsArray2.length; i++) {
            picArr.push(column1contentsArray2[i])
          }

          for (var i = 1; i < column1contentsArray3.length; i++) {
            desArr.push(column1contentsArray3[i])
          }

          for (var i = 1; i < column1contentsArray4.length; i++) {
            iconArr.push(column1contentsArray4[i])
          }

          for (var i = 1; i < column1contentsArray5.length; i++) {
            diviArr.push(column1contentsArray5[i])
          }

          for (var i = 1; i < column1contentsArray6.length; i++) {
            retailArr.push(column1contentsArray6[i])
          }

          for (var i = 1; i < column1contentsArray7.length; i++) {
            promoArr.push(column1contentsArray7[i])
          }

          for (var i = 1; i < column1contentsArray8.length; i++) {
            savingArr.push(column1contentsArray8[i])
          }
          for (var i = 1; i < column1contentsArray9.length; i++) {
            precentDiscountArr.push(column1contentsArray9[i])
          }

          for (var i = 1; i < column1contentsArray10.length; i++) {
            MoreIconsArr.push(column1contentsArray10[i])
          }

          // varable to set and control where to place the data in the page
          var counterFrame1 = 0

          var doc = app.activeDocument
          doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN
          var pageIdx = 0

          test()

          //if the product has another extenstion then return true
          function checkSibilngs(image) {
            matchingImages = []
            for (var i = 0; i <= allFiles.length; i++) {
              if (allFiles[i] != '../.DS_Store') {
                if (allFiles[i] !== undefined) {
                  if (allFiles[i].toString().indexOf(image) >= 0) {
                    matchingImages.push(allFiles[i])
                  } else {
                    continue
                  }
                } else {
                  continue
                }
              } else {
                continue
              }
            }

            if (matchingImages.length > 1) {
              return true
            } else {
              return false
            }
          }

          // if one product has many formats the function will check all the available formats
          // chose one of the available formats
          function renderOneSibling(matchingImages) {
            var countPSD = 0
            var countJPG = 0

            for (var i = 0; i < matchingImages.length; i++) {
              if (matchingImages[i].toString().indexOf('.png') >= 0) {
                countPSD = 0
                return matchingImages[i]
              } else {
                countPSD++
                continue
              }
            }

            if (countPSD > 0) {
              for (var i = 0; i < matchingImages.length; i++) {
                if (matchingImages[i].toString().indexOf('.psd') >= 0) {
                  countJPG = 0
                  return matchingImages[i]
                } else {
                  countJPG++
                  continue
                }
              }
            }
            if (countJPG > 0) {
              for (var i = 0; i < matchingImages.length; i++) {
                if (matchingImages[i].toString().indexOf('.jpg') >= 0) {
                  return matchingImages[i]
                } else {
                  var temp = matchingImages[i]
                  return matchingImages[i]
                }
              }
            }
          }

          // function to retrive the path of image(link) after check if the product has an image link
          // checking by taken th product code and loop throug the images folder and check if found it will return it

          function checkImage(image) {
            for (var i = 0; i <= allFiles.length; i++) {
              if (allFiles[i] != '../.DS_Store') {
                if (allFiles[i] !== undefined) {
                  if (allFiles[i].toString().indexOf(image) >= 0) {
                    //send each product to a function to check if this function has another diffrent extentions

                    var hasSublings = checkSibilngs(image)

                    if (hasSublings) {
                      // to pick the perfect extentions
                      return renderOneSibling(matchingImages)
                    } else {
                      return allFiles[i]
                    }
                  } else {
                    continue
                  }
                } else {
                  continue
                }
              } else {
                continue
              }
            }
          }
          // this function read the promotion type from the excel sheet and
          //combained it to one string and then search among the file of icones
          // of the combained promotion string match the url then retrive the url of the icon
          function checkIcon(
            icon,
            divi,
            retail,
            promo,
            saving,
            precentage,
            MoreIcons
          ) {
            if (icon.indexOf('B/G') >= 0) {
              placeMoreThanOneIcone(MoreIcons, divi, retail, promo, saving)
            } else if (icon.indexOf('2 @') >= 0) {
              placeMoreThanOneIcone(MoreIcons, divi, retail, promo, saving)
            } else if (icon.indexOf('Direct_Discount') >= 0) {
              if (icon.indexOf('%') >= 0) {
                placeNOExtraIconPresentage(icon, divi, precentage)
                // direct discount
                //alert(icon + " i'm the one WITH %")
              } else {
                placeMoreThanOneIcone(MoreIcons, divi, retail, promo, saving)
              }
            } else {
              iconePthe = icon + ' ' + divi

              for (var i = 0; i <= allFilesIcone.length; i++) {
                if (allFilesIcone[i] != '../.DS_Store') {
                  if (allFilesIcone[i] !== undefined) {
                    if (
                      decodeURI(allFilesIcone[i])
                        .toString()
                        .indexOf(iconePthe) >= 0
                    ) {
                      // IF LINK IS VALID AND FOUND
                      // alert( iconePthe + " NOT")
                      // alert (MoreIcons)

                      if (MoreIcons.length == 0) {
                        return allFilesIcone[i]
                      } else {
                        placeOneExtraIcon(icon, MoreIcons, divi, promo)
                      }
                    } else {
                      continue
                    }
                  } else {
                    continue
                  }
                } else {
                  continue
                }
              }
            }
          }

          function placeOneExtraIcon(icon, MoreIcons, divi, promo) {
            var icons = new Array()
            iconePthe = MoreIcons + ' ' + divi
            iconePthe1 = icon + ' ' + divi
            icons.push(iconePthe1)
            icons.push(iconePthe)
            for (var i = 0; i <= icons.length - 1; i++) {
              iconePth = icons[i]
              for (var j = 0; j <= allFilesIcone.length; j++) {
                if (allFilesIcone[j] != '../.DS_Store') {
                  if (allFilesIcone[j] !== undefined) {
                    if (
                      decodeURI(allFilesIcone[j])
                        .toString()
                        .indexOf(iconePth) >= 0
                    ) {
                      if (i > 0) {
                        z = arrayOfSize[counterFrame1]
                        z[0] = z[0] + 15
                        z[2] = z[2] + 15

                        // set the frame with the new geometric bound
                        rect2 = doc.pages[pageIdx].textFrames.add({
                          geometricBounds: z,
                        })
                        // placing the image and adding a function will resize the frame with the size of product
                        setEqualData(
                          allFilesIcone[j],
                          rect2,
                          counterFrame1,
                          divi,
                          promo,
                          pageIdx
                        )

                        rect2.fit(FitOptions.FRAME_TO_CONTENT)
                        rect2.fit(FitOptions.PROPORTIONALLY)
                        z = arrayOfSize[counterFrame1]
                      } else {
                        var rect2 = doc.pages[pageIdx].textFrames.add({
                          geometricBounds: arrayOfSize[counterFrame1],
                        })

                        // placing the image and adding a function will resize the frame with the size of product
                        rect2.place(allFilesIcone[j])
                        rect2.fit(FitOptions.FRAME_TO_CONTENT)
                        rect2.fit(FitOptions.PROPORTIONALLY)
                      }
                    } else {
                      continue
                    }
                  } else {
                    continue
                  }
                } else {
                  continue
                }
              }
            }
          }
          function placeNOExtraIconPresentage(icon, divi, precentage) {
            iconePthe = icon + ' ' + divi

            for (var j = 0; j <= allFilesIcone.length; j++) {
              if (allFilesIcone[j] != '../.DS_Store') {
                if (allFilesIcone[j] !== undefined) {
                  if (
                    decodeURI(allFilesIcone[j]).toString().indexOf(iconePthe) >=
                    0
                  ) {
                    var rectt2 = doc.pages[pageIdx].textFrames.add({
                      geometricBounds: arrayOfSize[counterFrame1],
                    })

                    setPrecentageData(
                      allFilesIcone[j],
                      rectt2,
                      counterFrame1,
                      divi,
                      precentage,
                      pageIdx
                    )
                    rectt2.fit(FitOptions.FRAME_TO_CONTENT)
                    rectt2.fit(FitOptions.PROPORTIONALLY)
                  } else {
                    continue
                  }
                } else {
                  continue
                }
              } else {
                continue
              }
            }
          }

          function placeMoreThanOneIcone(
            MoreIcons,
            divi,
            retail,
            promo,
            saving
          ) {
            if (MoreIcons.indexOf(',') >= 0) {
              var icons = new Array()
              var Tempicons = new Array()
              icons = MoreIcons.split(',')

              for (var i = 0; i <= icons.length - 1; i++) {
                iconePthe = icons[i] + ' ' + divi
                // alert ( iconePthe  + " " + i  + " " +icons.length);
                for (var j = 0; j <= allFilesIcone.length; j++) {
                  if (allFilesIcone[j] != '../.DS_Store') {
                    if (allFilesIcone[j] !== undefined) {
                      if (
                        decodeURI(allFilesIcone[j])
                          .toString()
                          .indexOf(iconePthe) >= 0
                      ) {
                        if (i > 0) {
                          l = arrayOfSize[counterFrame1]
                          l[1] = l[1] + 20
                          l[3] = l[3] + 20

                          // set the frame with the new geometric bound
                          var rectIcon = doc.pages[pageIdx].textFrames.add({
                            geometricBounds: l,
                          })

                          // placing the image and adding a function will resize the frame with the size of product
                          if (
                            decodeURI(allFilesIcone[j])
                              .toString()
                              .indexOf('Saving') >= 0
                          ) {
                            setSavingData(
                              allFilesIcone[j],
                              rectIcon,
                              counterFrame1,
                              divi,
                              saving,
                              pageIdx
                            )
                          }
                          if (
                            decodeURI(allFilesIcone[j])
                              .toString()
                              .indexOf('Direct_Discount') >= 0
                          ) {
                            // && (decodeURI(allFilesIcone[j]).toString().indexOf("&") <= 0)
                            //  alert(allFilesIcone[j])
                            setDirectDiscountData(
                              allFilesIcone[j],
                              rectIcon,
                              counterFrame1,
                              divi,
                              retail,
                              promo,
                              pageIdx
                            )
                          } else {
                            rectIcon.place(allFilesIcone[j])
                          }
                          rectIcon.fit(FitOptions.FRAME_TO_CONTENT)
                          rectIcon.fit(FitOptions.PROPORTIONALLY)

                          // rectIcon.select(SelectionOptions.ADD_TO)
                          l = arrayOfSize[counterFrame1]
                        } else {
                          // set the frame with the new geometric bound
                          var rectIcon = doc.pages[pageIdx].textFrames.add({
                            geometricBounds: arrayOfSize[counterFrame1],
                          })

                          // placing the image and adding a function will resize the frame with the size of product
                          if (
                            decodeURI(allFilesIcone[j])
                              .toString()
                              .indexOf('Value') >= 0
                          ) {
                            setValueData(
                              allFilesIcone[j],
                              rectIcon,
                              counterFrame1,
                              divi,
                              promo,
                              pageIdx
                            )
                          } else {
                            rectIcon.place(allFilesIcone[j])
                          }
                          rectIcon.fit(FitOptions.FRAME_TO_CONTENT)
                          rectIcon.fit(FitOptions.PROPORTIONALLY)
                        }
                      } else {
                        continue
                      }
                    } else {
                      continue
                    }
                  } else {
                    continue
                  }
                }
              }
            }
          }
          // this function to vaildate if the user inputs
          // The correct input should be numbers
          //   function checkPopup(slot1) {
          //     var numbers = /^[1-9]+$/

          //     do {
          //       if (slot1 == '' || slot1 == null) {
          //         alert('There is NO value please enter value')
          //         slot1 = prompt(
          //           '',
          //           '',
          //           ' How many product you want to insert in this group?'
          //         )
          //       }
          //       if (slot1 == 0) {
          //         alert('Please enter number LARGER than 0')
          //         slot1 = prompt(
          //           '',
          //           '',
          //           ' How many product you want to insert in this group?'
          //         )
          //       }
          //       if (slot1 > idArr.length) {
          //         alert(
          //           'please enter number SMALLER than the range of the excel sheet'
          //         )
          //         slot1 = prompt(
          //           '',
          //           '',
          //           ' How many product you want to insert in this group?'
          //         )
          //       }
          //       if (!slot1.match(numbers)) {
          //         alert('please enter ONLY numbers')
          //         slot1 = prompt(
          //           '',
          //           '',
          //           ' How many product you want to insert in this group?'
          //         )
          //       }
          //       if (slot1.match(numbers)) {
          //         return slot1
          //       }
          //     } while (isNaN(slot1) || slot1 == null || slot1 < 0)
          //   }

          // The main process start after calling it and activate above
          function test() {
            // slot1 = prompt(
            //   '',
            //   '',
            //   ' How many product you want to insert in this group?'
            // )
            // testSlot = checkPopup(slot1)

            for (var i = 1; i < idArr.length; i++) {
              // loop to compare

              if (i > 1) {
                //to compare between each data column so do not duplacte
                //the result and skip the data we already read and procesdd
                if (idArr[i] == idArr[i - 1]) {
                  continue
                }
              }

              // second loop to be compared with
              for (var j = 0; j < idArr.length; j++) {
                //if the group found it will push all the other group info
                if (idArr[i] == idArr[j]) {
                  // pusht matched data
                  tempArr.push(idArr[j])
                  tempArr2.push(picArr[j])
                  tempArr3.push(desArr[j])
                  tempArr4.push(iconArr[j])
                  tempArr5.push(diviArr[j])
                  tempArr6.push(retailArr[j])
                  tempArr7.push(promoArr[j])
                  tempArr8.push(savingArr[j])
                  tempArr9.push(precentDiscountArr[j])
                  tempArr10.push(MoreIconsArr[j])
                }
              }

              //write the data in each frame

              // the size of the array will be resitted by the user
              // lets say in the excel we have a group called "10203_1" and in this group
              // we have 19 products and the user enter 4, the user will specify the maximum number
              //of products for all the groups and then based on that the system will stop loop until 4

              //column1contentsArray11[i] array contains the number of Products for each product
              //WARNING! Original logic proceeds after passing the respective number of products
              for (var s = 0; s <= column1contentsArray11[i] - 1; s++) {
                // loop through the array of products

                // send each product to a function that will check if the image exists or not.
                xTest = checkImage(tempArr2[s])
                // alert(tempArr2[s])

                // if exists but it's undefied the system will ignore it.

                if (xTest !== undefined) {
                  // Now since we want the product images to be side by side aligned
                  // We have to postion the first image (0) in the same predeaclerd postion in the Arrayofsize
                  // and the other images (1)  and (2) for each time the system will postion the image it will add
                  // 15mm from the left and the right, so after postioning the images we going to redeclear the
                  //Arrayofsize to reset the old postions and start over correctly without extra spaces
                  if (s > 0) {
                    // save the coordnate in varabile so we can access the geometric data and manuplate it
                    a = arrayOfSize[counterFrame1]
                    a[1] = a[1] + 20
                    a[3] = a[3] + 20

                    // set the frame with the new geometric bound
                    var rect = doc.pages[pageIdx].textFrames.add({
                      geometricBounds: a,
                    })

                    // placing the image and adding a function will resize the frame with the size of product
                    rect.place(xTest)
                    rect.fit(FitOptions.PROPORTIONALLY)
                    rect.fit(FitOptions.FRAME_TO_CONTENT)
                    rect.select(SelectionOptions.ADD_TO)
                  } else {
                    // defualt setting
                    var rect = doc.pages[pageIdx].textFrames.add({
                      geometricBounds: arrayOfSize[counterFrame1],
                    })

                    rect.place(xTest)
                    rect.fit(FitOptions.PROPORTIONALLY)
                    rect.fit(FitOptions.FRAME_TO_CONTENT)
                    rect.select(SelectionOptions.ADD_TO)
                  }
                } else {
                  continue
                }
              }

              // Redeclearing the array to reset postions
              // because after the aligemnt for each group the postioning coordnate will be affected

              arrayOfSize = [
                [18, 18, 68, 74.54],
                [18, 120, 68, 181.84],
                [110, 12.7, 160.84, 74.54],
                [110, 120, 160, 181.84],
                [206, 12.7, 256.84, 74.54],
                [206, 120, 256.84, 181.84],
              ]
              //placing the description
              var rect1 = doc.pages[pageIdx].textFrames.add({
                geometricBounds: arrayOfdescrptionframe[counterFrame1],
              })
              rect1.contents = tempArr3[0]
              //changing the type of font for the description and resize it
              rect1.texts[0].appliedFont = 'Nahdi	Black'
              rect1.texts[0].pointSize = 7
              rect1.texts[0].parentStory.justification =
                Justification.CENTER_ALIGN
              rect1.texts[0].fillColor = color
              // to fit the icons in the frame with the old geometric boundryes

              // placing the icones

              // to check weathe the icon is valid or not
              xIconeTest = checkIcon(
                tempArr4[0],
                tempArr5[0],
                tempArr6[0],
                tempArr7[0],
                tempArr8[0],
                tempArr9[0],
                tempArr10[0]
              )
              if (xIconeTest !== undefined) {
                if (xIconeTest !== ' ') {
                  var rect2 = doc.pages[pageIdx].textFrames.add({
                    geometricBounds: arrayOfSize[counterFrame1],
                  })
                  rect2.place(xIconeTest)
                  rect2.fit(FitOptions.FRAME_TO_CONTENT)
                  rect2.fit(FitOptions.PROPORTIONALLY)
                }
              }
              // if the icon is invalid it will not render it

              // emptying the temp arrayes to read a new group
              tempArr = []
              tempArr2 = []
              tempArr3 = []
              tempArr4 = []
              tempArr5 = []
              tempArr6 = []
              tempArr7 = []
              tempArr8 = []
              tempArr9 = []
              tempArr10 = []

              // set the margin for each page in the document
              doc.pages[pageIdx].marginPreferences.properties = {
                top: '18 mm',
                left: '10 mm',
                right: '10 mm',
                bottom: '13 mm',
              }

              counterFrame1++

              //to make sure if there is more than 6 frames in spcific page add anothr page and create frames on it.
              if (counterFrame1 % 6 == 0) {
                counterFrame1 = 0
                //Add a page
                doc.pages.add()
                pageIdx++
              }
            }
          }
        } catch (e) {
          alert(e)
          doc.close(1852776480)
        }
      } catch (e) {
        alert(e)
        doc.close(1852776480)
      }
    }

    function setPrecentageData(icone, rect2, frame, division, Saving, pageIdx) {
      rect2.place(icone)

      var beuaty0 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 40, 52, 55],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG0 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 40, 52, 55],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby0 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 40, 52, 55],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS0 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 40, 52, 55],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 1 [12.7, 120, 90, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty1 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 144, 52, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG1 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 144, 52, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby1 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 144, 52, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS1 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 144, 52, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 2 [110, 12.7, 190, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty2 = function () {
        //	[141, 53, 146, 65]  [127, 43, 133, 56]   [136, 53, 141, 65]

        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 146, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG2 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 146, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby2 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 146, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS2 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 146, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 3 [110, 120, 190, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty3 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 144, 146, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG3 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 144, 146, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby3 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 144, 146, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS3 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 144, 146, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 4 [206, 12.7, 284, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty4 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 37, 245, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG4 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 37, 245, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby4 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 37, 245, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS
      var WELLNESS4 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 37, 245, 54],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 5 [206, 120, 284, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty5 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 144, 245, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG5 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 144, 245, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby5 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 144, 245, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS5 = function () {
        var rectSaving = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [235, 144, 245, 160],
        })

        rectSaving.contents = '%' + Saving

        rectSaving.texts[0].appliedFont = 'Variable	Black'

        rectSaving.texts[0].pointSize = 17

        rectSaving.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectSaving.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      // The diction that contain all the functions
      // the keys are the postions and the values are the divtions
      // lets say, we have postion 1 division MOM AND BABY then the function momandbaby1 will called and activated
      var dict = {
        0: {
          BEAUTY: beuaty0,
          FMCG: FMCG0,
          'MOM AND BABY': momandbaby0,
          WELLNESS: WELLNESS0,
        },
        1: {
          BEAUTY: beuaty1,
          FMCG: FMCG1,
          'MOM AND BABY': momandbaby1,
          WELLNESS: WELLNESS1,
        },
        2: {
          BEAUTY: beuaty2,
          FMCG: FMCG2,
          'MOM AND BABY': momandbaby2,
          WELLNESS: WELLNESS2,
        },
        3: {
          BEAUTY: beuaty3,
          FMCG: FMCG3,
          'MOM AND BABY': momandbaby3,
          WELLNESS: WELLNESS3,
        },
        4: {
          BEAUTY: beuaty4,
          FMCG: FMCG4,
          'MOM AND BABY': momandbaby4,
          WELLNESS: WELLNESS4,
        },
        5: {
          BEAUTY: beuaty5,
          FMCG: FMCG5,
          'MOM AND BABY': momandbaby5,
          WELLNESS: WELLNESS5,
        },
      }

      var tex = division
      // loop through the dictionry
      for (var key in dict) {
        // key = postion (frame)
        if (key == frame) {
          for (var x in dict[key]) {
            // x = division (temp5[0])
            if (x.indexOf(tex) >= 0) {
              // calling the function
              dict[key][tex]()
            }
          }
        }
      }
    }
    function setEqualData(icone, rect2, frame, division, Promo, pageIdx) {
      rect2.place(icone)

      var beuaty0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 37, 65, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 37, 65, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 37, 65, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 37, 65, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 1 [12.7, 120, 90, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 143, 65, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 143, 65, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 143, 65, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [60, 143, 65, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 2 [110, 12.7, 190, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty2 = function () {
        //	[141, 53, 146, 65]  [127, 43, 133, 56]   [136, 53, 141, 65]
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [155, 35, 161, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG2 = function () {
        //	[141, 53, 146, 65]  [127, 43, 133, 56]   [136, 53, 141, 65]
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [155, 35, 161, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby2 = function () {
        //	[141, 53, 146, 65]  [127, 43, 133, 56]   [136, 53, 141, 65]
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [155, 35, 161, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS2 = function () {
        //	[141, 53, 146, 65]  [127, 43, 133, 56]   [136, 53, 141, 65]
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [155, 35, 161, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 3 [110, 120, 190, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [154, 142, 160, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [154, 142, 160, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [154, 142, 160, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [154, 142, 160, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 4 [206, 12.7, 284, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 35, 256, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 35, 256, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 35, 256, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS
      var WELLNESS4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 35, 256, 50],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 5 [206, 120, 284, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 142, 256, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 142, 256, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 142, 256, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [251, 142, 256, 157],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 6

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      var dict = {
        0: {
          BEAUTY: beuaty0,
          FMCG: FMCG0,
          'MOM AND BABY': momandbaby0,
          WELLNESS: WELLNESS0,
        },
        1: {
          BEAUTY: beuaty1,
          FMCG: FMCG1,
          'MOM AND BABY': momandbaby1,
          WELLNESS: WELLNESS1,
        },
        2: {
          BEAUTY: beuaty2,
          FMCG: FMCG2,
          'MOM AND BABY': momandbaby2,
          WELLNESS: WELLNESS2,
        },
        3: {
          BEAUTY: beuaty3,
          FMCG: FMCG3,
          'MOM AND BABY': momandbaby3,
          WELLNESS: WELLNESS3,
        },
        4: {
          BEAUTY: beuaty4,
          FMCG: FMCG4,
          'MOM AND BABY': momandbaby4,
          WELLNESS: WELLNESS4,
        },
        5: {
          BEAUTY: beuaty5,
          FMCG: FMCG5,
          'MOM AND BABY': momandbaby5,
          WELLNESS: WELLNESS5,
        },
      }

      var tex = division
      // loop through the dictionry
      for (var key in dict) {
        // key = postion (frame)
        if (key == frame) {
          for (var x in dict[key]) {
            // x = division (temp5[0])
            if (x.indexOf(tex) >= 0) {
              // calling the function
              dict[key][tex]()
            }
          }
        }
      }
    }

    function setValueData(icone, rect2, frame, division, Promo, pageIdx) {
      rect2.place(icone)

      var beuaty0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 39, 49, 54],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 39, 49, 54],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 39, 49, 54],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 39, 49, 54],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 1 [12.7, 120, 90, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 143, 49, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 143, 49, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 143, 49, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [44, 143, 49, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 2 [110, 12.7, 190, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 144, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 144, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 144, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 37, 144, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 3 [110, 120, 190, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 143, 144, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 143, 144, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 143, 144, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [138, 143, 144, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 4 [206, 12.7, 284, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 37, 240, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 37, 240, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 37, 240, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS
      var WELLNESS4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 37, 240, 53],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 5 [206, 120, 284, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 143, 240, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 143, 240, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 143, 240, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [234, 143, 240, 160],
        })

        rectPromo.contents = Promo.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 11

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 9

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      var dict = {
        0: {
          BEAUTY: beuaty0,
          FMCG: FMCG0,
          'MOM AND BABY': momandbaby0,
          WELLNESS: WELLNESS0,
        },
        1: {
          BEAUTY: beuaty1,
          FMCG: FMCG1,
          'MOM AND BABY': momandbaby1,
          WELLNESS: WELLNESS1,
        },
        2: {
          BEAUTY: beuaty2,
          FMCG: FMCG2,
          'MOM AND BABY': momandbaby2,
          WELLNESS: WELLNESS2,
        },
        3: {
          BEAUTY: beuaty3,
          FMCG: FMCG3,
          'MOM AND BABY': momandbaby3,
          WELLNESS: WELLNESS3,
        },
        4: {
          BEAUTY: beuaty4,
          FMCG: FMCG4,
          'MOM AND BABY': momandbaby4,
          WELLNESS: WELLNESS4,
        },
        5: {
          BEAUTY: beuaty5,
          FMCG: FMCG5,
          'MOM AND BABY': momandbaby5,
          WELLNESS: WELLNESS5,
        },
      }

      var tex = division
      // loop through the dictionry
      for (var key in dict) {
        // key = postion (frame)
        if (key == frame) {
          for (var x in dict[key]) {
            // x = division (temp5[0])
            if (x.indexOf(tex) >= 0) {
              // calling the function
              dict[key][tex]()
            }
          }
        }
      }
    }
    function setSavingData(icone, rect2, frame, division, savingVal, pageIdx) {
      rect2.place(icone)

      var beuaty0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 60, 50, 74],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 60, 50, 74],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 60, 50, 74],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 60, 50, 74],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 1 [12.7, 120, 90, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 163, 50, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 163, 50, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 163, 50, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 163, 50, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 2 [110, 12.7, 190, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 57, 145, 72],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 57, 145, 72],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 57, 145, 72],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 57, 145, 72],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 3 [110, 120, 190, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 163, 145, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 163, 145, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 163, 145, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 163, 145, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 4 [206, 12.7, 284, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 57, 241, 73],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 57, 241, 73],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 57, 241, 73],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS
      var WELLNESS4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 57, 241, 73],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 5 [206, 120, 284, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 163, 241, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 163, 241, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 163, 241, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 163, 241, 180],
        })

        rectPromo.contents = savingVal.toString()

        rectPromo.texts[0].appliedFont = 'Variable	Black'

        rectPromo.texts[0].pointSize = 10

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      var dict = {
        0: {
          BEAUTY: beuaty0,
          FMCG: FMCG0,
          'MOM AND BABY': momandbaby0,
          WELLNESS: WELLNESS0,
        },
        1: {
          BEAUTY: beuaty1,
          FMCG: FMCG1,
          'MOM AND BABY': momandbaby1,
          WELLNESS: WELLNESS1,
        },
        2: {
          BEAUTY: beuaty2,
          FMCG: FMCG2,
          'MOM AND BABY': momandbaby2,
          WELLNESS: WELLNESS2,
        },
        3: {
          BEAUTY: beuaty3,
          FMCG: FMCG3,
          'MOM AND BABY': momandbaby3,
          WELLNESS: WELLNESS3,
        },
        4: {
          BEAUTY: beuaty4,
          FMCG: FMCG4,
          'MOM AND BABY': momandbaby4,
          WELLNESS: WELLNESS4,
        },
        5: {
          BEAUTY: beuaty5,
          FMCG: FMCG5,
          'MOM AND BABY': momandbaby5,
          WELLNESS: WELLNESS5,
        },
      }

      var tex = division
      // loop through the dictionry
      for (var key in dict) {
        // key = postion (frame)
        if (key == frame) {
          for (var x in dict[key]) {
            // x = division (temp5[0])
            if (x.indexOf(tex) >= 0) {
              // calling the function
              dict[key][tex]()
            }
          }
        }
      }
    }
    function setDirectDiscountData(
      icone,
      rect2,
      frame,
      division,
      retaill,
      Promo,
      pageIdx
    ) {
      rect2.place(icone)
      var beuaty0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 80, 50, 94],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 80, 45, 94],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 80, 50, 94],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 80, 45, 94],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 80, 50, 94],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 80, 45, 94],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS0 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 80, 50, 94],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 80, 45, 94],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 1 [12.7, 120, 90, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 183, 50, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 183, 45, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 183, 50, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 183, 45, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 183, 50, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 183, 45, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS1 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [45, 183, 50, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [40, 183, 45, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 2 [110, 12.7, 190, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */
      var beuaty2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 77, 145, 92],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 77, 140, 92],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 77, 145, 92],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 77, 140, 92],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 77, 145, 92],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 77, 140, 92],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS2 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 77, 145, 92],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 77, 140, 92],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 3 [110, 120, 190, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 183, 145, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 183, 140, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 183, 145, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 183, 140, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 183, 145, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 183, 140, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** WELLNESS

      var WELLNESS3 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [140, 183, 145, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [135, 183, 140, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 4 [206, 12.7, 284, 90]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 77, 241, 93],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 77, 236, 93],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 77, 241, 93],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 77, 236, 93],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY
      var momandbaby4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 77, 241, 93],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 77, 236, 93],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS
      var WELLNESS4 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 77, 241, 93],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 77, 236, 93],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //00000000000000000000000000000000000000000000000000000000000000000000000000000
      // POSTION 5 [206, 120, 284, 197.3]
      //0000000000000000000000000000000000000000000000000000000000000000000000000000
      //******************************************************BEUYT ICON */

      var beuaty5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 183, 241, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 183, 236, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }
      //****************************************************** FMCG ICONE */

      var FMCG5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 183, 241, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 183, 236, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** MOM AND BABY

      var momandbaby5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 183, 241, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 183, 236, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      //****************************************************** WELLNESS

      var WELLNESS5 = function () {
        var rectPromo = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [236, 183, 241, 200],
        })

        var rectRetail = doc.pages[pageIdx].textFrames.add({
          geometricBounds: [231, 183, 236, 200],
        })

        rectPromo.contents = Promo.toString()
        rectRetail.contents = retaill.toString()
        rectPromo.texts[0].appliedFont = 'Variable	Black'
        rectRetail.texts[0].appliedFont = 'Variable	Black'
        rectPromo.texts[0].pointSize = 10
        rectRetail.texts[0].pointSize = 7

        app.findGrepPreferences.findWhat = '\\d+\\.\\K\\d+'

        var promo = rectPromo.findGrep()
        if (promo.length) promo[0].pointSize = 8

        var retail = rectRetail.findGrep()
        if (retail.length) retail[0].pointSize = 5

        rectPromo.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectPromo.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN

        rectRetail.texts[0].fillColor = app.documents[0].swatches.itemByName(
          'Paper'
        )
        rectRetail.texts[0].parentStory.justification =
          Justification.CENTER_ALIGN
      }

      var dict = {
        0: {
          BEAUTY: beuaty0,
          FMCG: FMCG0,
          'MOM AND BABY': momandbaby0,
          WELLNESS: WELLNESS0,
        },
        1: {
          BEAUTY: beuaty1,
          FMCG: FMCG1,
          'MOM AND BABY': momandbaby1,
          WELLNESS: WELLNESS1,
        },
        2: {
          BEAUTY: beuaty2,
          FMCG: FMCG2,
          'MOM AND BABY': momandbaby2,
          WELLNESS: WELLNESS2,
        },
        3: {
          BEAUTY: beuaty3,
          FMCG: FMCG3,
          'MOM AND BABY': momandbaby3,
          WELLNESS: WELLNESS3,
        },
        4: {
          BEAUTY: beuaty4,
          FMCG: FMCG4,
          'MOM AND BABY': momandbaby4,
          WELLNESS: WELLNESS4,
        },
        5: {
          BEAUTY: beuaty5,
          FMCG: FMCG5,
          'MOM AND BABY': momandbaby5,
          WELLNESS: WELLNESS5,
        },
      }

      var tex = division
      // loop through the dictionry
      for (var key in dict) {
        // key = postion (frame)
        if (key == frame) {
          for (var x in dict[key]) {
            // x = division (temp5[0])
            if (x.indexOf(tex) >= 0) {
              // calling the function
              dict[key][tex]()
            }
          }
        }
      }
    }

    app.activeWindow.viewDisplaySetting = ViewDisplaySettings.HIGH_QUALITY
    doc.viewPreferences.properties = currViewPrefs
  }

  // calling the function && let the user set the desired excel file
  app.doScript(
    "start(File.openDialog('select file', '*.*'))",
    ScriptLanguage.JAVASCRIPT,
    undefined,
    UndoModes.ENTIRE_SCRIPT,
    'test'
  )
}

(ns haydn.core
  (:import [java.io File FileOutputStream]
           [java.awt Color]
           [org.apache.poi.common.usermodel HyperlinkType]
           [org.apache.poi.ss.util CellRangeAddress]
           (org.apache.poi.ss.usermodel CellStyle
                                        CellType
                                        CreationHelper
                                        FillPatternType
                                        HorizontalAlignment
                                        IndexedColors)
           (org.apache.poi.xssf.usermodel XSSFWorkbook
                                          XSSFSheet
                                          XSSFFont
                                          XSSFColor
                                          DefaultIndexedColorMap
                                          XSSFHyperlink
                                          TextAlign
                                          XSSFRow)
           [org.apache.poi.xssf.usermodel XSSFWorkbook]))

;;global workbook
(def workbook (new XSSFWorkbook))

(defn tr? [element]
  (cond
    (not (keyword? element)) false
    (nil? (re-find #"^tr" (name element))) false
    :else true))

(defn th? [element]
  (cond
    (not (keyword? element)) false
    (nil? (re-find #"^th" (name element))) false
    :else true))

(defmulti parse-expr (fn [form obj]
                       (cond
                         (= :wb (first form)) :wb
                         (= :table (first form)) :table
                         (tr? (first form)) :tr
                         (= :td (first form)) :td
                         (= :tbody (first form)) :tbody
                         (= :thead (first form)) :thead
                         (= :tfoot (first form)) :tfoot
                         (= :a (first form)) :a
                         (nil? (first form)) :nil
                         (= 'for (first form)) :for
                         (th? (first form)) :th
                         (map? form) :map
                         (seq? form) :seq)))

(defn row-empty? [row]
  (cond
    (nil? row) true
    (<= (.getLastCellNum row) 0) true
    :else false))

(defmethod parse-expr :nil
  [form obj])

(defmethod parse-expr :tbody
  [form obj]
  (parse-expr (rest form) obj))

(defmethod parse-expr :a
  [form obj]
  (let [url (:href (second form))
        text (nth form 2)
        create-helper (.getCreationHelper workbook)
        link (.createHyperlink create-helper HyperlinkType/URL)]
    (.setCellValue obj text)
    (.setAddress link url)
    (.setHyperlink obj link)))

(defmethod parse-expr :thead
  [form obj]
  (parse-expr (rest form) obj))

(defmethod parse-expr :tfoot
  [form obj]
  (parse-expr (rest form) obj))

(defmethod parse-expr :for
  [form obj]
  (parse-expr (eval form) obj))

(defmethod parse-expr :wb
  [form obj]
  (parse-expr (rest form) obj))

(defmethod parse-expr :map
  [form obj]
  (if (= "org.apache.poi.xssf.usermodel.XSSFCell"
         (.getName (type obj))) (cond
                                  (contains? form :colspan)
                                  (do
                                    (.setAlignment (.getCellStyle obj) HorizontalAlignment/CENTER)
                                    (.setCellStyle obj (.getCellStyle obj))
                                    (.addMergedRegion (.getSheet obj) (new CellRangeAddress
                                                                           (.getRowNum (.getRow obj));;firstRow
                                                                           (.getRowNum (.getRow obj));;lastRow
                                                                           (.getColumnIndex obj)
                                                                           (dec (Integer. (:colspan form))))))))
  (if (= "org.apache.poi.xssf.usermodel.XSSFRow"
         (.getName (type obj)))
    (cond
      (contains? form :background-color)
      (let [style (.createCellStyle workbook)
            background-color (:background-color form)]
        (try
          (.setFillBackgroundColor style (.getIndex (IndexedColors/valueOf (.toUpperCase background-color))))
          (catch IllegalArgumentException e
            ;;not an indexed color, let's assume it's hex
            (.setFillForegroundColor style
                                     (new XSSFColor (Color/decode background-color)
                                          (new DefaultIndexedColorMap)))))
        (.setFillPattern style FillPatternType/SOLID_FOREGROUND)
        (.setRowStyle obj style)))))

(defmethod parse-expr :seq
  [form obj]
  (parse-expr (first form) obj)
  (if (seq (rest form))
    (parse-expr (rest form) obj)))

(defmethod parse-expr :table
  [form obj]
  (if (not (map? (second form)))
    (throw (Exception. "Title is required")))
  (let [title (:title (second form))
        sheet (.createSheet obj title)
        tr-list (rest (rest form))]
    (dotimes [n (count tr-list)]
      (parse-expr (nth tr-list n) sheet))))

(defmethod parse-expr :tr
  [form obj]
  (let [row (.createRow obj
                        (if (row-empty? (.getRow obj (.getLastRowNum obj)))
                          (.getLastRowNum obj)
                          (inc (.getLastRowNum obj))))
        td-list (rest form)]
    (dotimes [n (count td-list)]
      (parse-expr (nth td-list n) row))
    (dotimes [i (.getPhysicalNumberOfCells (.getRow obj 0))]
      (.autoSizeColumn obj i))))

(defmethod parse-expr :td
  [form obj]
  (let [cell (.createCell obj
                          (if (= -1 (.getLastCellNum obj))
                            0
                            (Integer. (str (.getLastCellNum obj)))))]
    (cond
      (vector? (first (rest form))) (parse-expr (first (rest form)) cell)
      (map? (first (rest form))) (let [value (second (rest form))]
                                   (parse-expr (first (rest form)) cell)
                                   (.setCellValue cell (second (rest form))))
      (integer? (first (rest form))) (.setCellValue
                                      cell
                                      (Double. (str (first (rest form)))))
      :else
      (.setCellValue cell (first (rest form))))))

(defmethod parse-expr :th
  [form row]
  (let [cell (.createCell row
                          (if (= -1 (.getLastCellNum row))
                            0
                            (Integer. (str (.getLastCellNum row)))))
        font (.createFont workbook)
        style (.getRowStyle row)]
    (.setBold font true)
    (.setFont style font)
    (.setAlignment style HorizontalAlignment/CENTER)
    (.setCellStyle cell style)
    (cond
      (integer? (first (rest form))) (.setCellValue
                                      cell
                                      (Double. (str (first (rest form)))))
      :else
      (.setCellValue cell (first (rest form))))))

(defmacro excel [form file]
  `(let [out# (new FileOutputStream (new File ~file))]
     (parse-expr '~form workbook)
     (.write workbook out#)
     (.close out#)))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  (excel [:wb
          [:table {:title "Test"}
           [:thead
            [:tr {:background-color "#8DBDD8"}
             [:th "President"]
             [:th "Born"]
             [:th "Died"]
             [:th "Wiki"]]]
           [:tbody
            [:tr [:td "Abraham Lincoln"]
             [:td "1809"]
             [:td "1865"]
             [:td
              [:a {:href "https://en.wikipedia.org/wiki/Abraham_Lincoln"} "Bio"]]]
            [:tr
             [:td "Andrew Johnson"]
             [:td "1808"]
             [:td "1875"]
             [:td [:a {:href "https://en.wikipedia.org/wiki/Andrew_Johnson"} "Bio"]]]
            [:tr
             [:td "Ulysses S. Grant"]
             [:td "1822"]
             [:td "1885"]
             [:td
              [:a {:href "https://en.wikipedia.org/wiki/Ulysses_S._Grant"} "Bio"]]]
            [:tr
             [:td "Rutherford B. Hayes"]
             [:td "1822"]
             [:td "1893"]
             [:td [:a {:href "https://en.wikipedia.org/wiki/Rutherford_B._Hayes"} "Bio"]]]
            ]
           [:tfoot
            [:tr [:td {:colspan "4"} "Reconstruction Presidents."]]]
           ]
          [:table {:title "First Sheet"}
           [:tr [:td "A"] [:td "B"] [:td "C"] [:td "D"]]
           [:tr [:td "E"] [:td "F"] [:td "G"] [:td "H"]]
           [:tr [:td "I"] [:td "J"] [:td "K"] [:td "L"]]]
          [:table {:title "Second Sheet"}
           [:tbody
            [:tr [:td "1"] [:td "2"] [:td "3"]]
            [:tr [:td "4"] [:td "5"] [:td "6"]]
            [:tr [:td "7"] [:td "8"] [:td "9"]]]]
          [:table {:title "For Test"}
           (for [x (range 5)]
             [:tr [:td x]])]]
         "haydn.xslx"))

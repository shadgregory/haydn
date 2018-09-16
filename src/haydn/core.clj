(ns haydn.core
  (:import [java.io File FileOutputStream]
           [org.apache.poi.xssf.usermodel XSSFWorkbook]))

(defmulti parse-expr (fn [form obj]
                       (cond
                         (= :wb (first form)) :wb
                         (= :table (first form)) :table
                         (= :tr (first form)) :tr
                         (= :td (first form)) :td
                         (= :tbody (first form)) :tbody
                         (= 'for (first form)) :for
                         (seq? form) :seq)))

(defn row-empty? [row]
  (cond
    (nil? row) true
    (<= (.getLastCellNum row) 0) true
    :else false))

(defmethod parse-expr :tbody
  [form obj]
  (parse-expr (rest form) obj))

(defmethod parse-expr :for
  [form obj]
  (parse-expr (eval form) obj))

(defmethod parse-expr :wb
  [form obj]
  (parse-expr (rest form) obj))

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
      (parse-expr (nth td-list n) row))))

(defmethod parse-expr :td
  [form obj]
  (let [cell (.createCell obj
                          (if (= -1 (.getLastCellNum obj))
                            0
                            (Integer. (str (.getLastCellNum obj)))))]
    (cond
      (integer? (first (rest form))) (.setCellValue
                                      cell
                                      (Double. (str (first (rest form)))))
      :else
      (.setCellValue cell (first (rest form))))))

(defmacro excel [form file]
  `(let [out# (new FileOutputStream (new File ~file))
         wb# (new XSSFWorkbook)]
     (parse-expr '~form wb#)
     (.write wb# out#)
     (.close out#)))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  (excel [:wb
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

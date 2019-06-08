(defproject haydn "0.1.0"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "GNU LESSER GENERAL PUBLIC LICENSE"}
  :dependencies [[org.clojure/clojure "1.10.1"]
                 [com.taoensso/tufte "2.0.1"]
                 [org.apache.poi/poi "4.1.0"]
                 [org.apache.poi/poi-ooxml "4.1.0"]]
  :local-repo "local-m2"
  :main ^:skip-aot haydn.core
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all}})

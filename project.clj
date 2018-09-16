(defproject haydn "0.1.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "GNU LESSER GENERAL PUBLIC LICENSE"}
  :dependencies [[org.clojure/clojure "1.9.0"]
                 [org.apache.poi/poi "4.0.0"]
                 [org.apache.poi/poi-ooxml "4.0.0"]]
  :local-repo "local-m2"
  :main ^:skip-aot haydn.core
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all}})

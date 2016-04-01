val poi = Seq(
  "org.scala-lang.modules" %% "scala-xml" % "1.0.3",
  "org.apache.poi" % "poi" % "3.14",
  "org.apache.poi" % "poi-ooxml" % "3.14"
)

lazy val root = (project in file(".")).
  settings(
    name := "excel2xml",
    version := "1.0",
    scalaVersion := "2.11.7",
    libraryDependencies ++= poi
  )


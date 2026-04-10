import com.kms.katalon.core.annotation.BeforeTestCase
import com.kms.katalon.core.context.TestCaseContext
import java.io.File

class TempCleaner {

    @BeforeTestCase
    def cleanBeforeRun(TestCaseContext testCaseContext) {

        println("===== TEMP CLEANUP START =====")
        println("Running before: " + testCaseContext.getTestCaseId())

        File tempDir = new File(System.getProperty("java.io.tmpdir"))

        tempDir.listFiles()?.each { file ->
            if (file.isDirectory() &&
                (
                    file.name.startsWith("katalon-cft") ||
                    file.name.startsWith("katalon-clean") ||
                    file.name.startsWith("scoped_dir")
                )
            ) {
                try {
                    file.deleteDir()
                    println("Deleted temp folder: " + file.absolutePath)
                } catch (Exception e) {
                    println("Skipped temp folder: " + file.absolutePath)
                }
            }
        }

        println("===== TEMP CLEANUP END =====")
    }
}
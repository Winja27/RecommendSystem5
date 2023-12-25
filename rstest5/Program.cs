// <SnippetUsingStatements>
using System;
using System.Data;
using System.IO;
using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.ML.Trainers;
using OfficeOpenXml;
// </SnippetUsingStatements>

namespace MovieRecommendation
{
    class Program
    {
        static void Main(string[] args)
        {

            // 创建 MLContext 对象，该对象在整个模型创建工作流中共享
            // <SnippetMLContext>
            MLContext mlContext = new MLContext();
            // </SnippetMLContext>

            // 载入数据
            // <SnippetLoadDataMain>
            (IDataView trainingDataView, IDataView testDataView) = LoadData(mlContext);
            // </SnippetLoadDataMain>

            // 构建并训练模型
            // <SnippetBuildTrainModelMain>
            ITransformer model = BuildAndTrainModel(mlContext, trainingDataView);
            // </SnippetBuildTrainModelMain>

            // 评估模型质量
            // <SnippetEvaluateModelMain>
            EvaluateModelAndWriteToExcel(mlContext, testDataView, model, "/metrics.xlsx");
            // </SnippetEvaluateModelMain>

            // 使用模型进行单个预测（一行数据）
            // <SnippetUseModelMain>
            // UseModelForSinglePrediction(mlContext, model);
            // </SnippetUseModelMain>

            // 保存模型
            // <SnippetSaveModelMain>
            // SaveModel(mlContext, trainingDataView.Schema, model);
            // </SnippetSaveModelMain>
        }

        // 载入数据
        public static (IDataView training, IDataView test) LoadData(MLContext mlContext)
        {
            // 使用数据路径载入训练和测试数据集
            // <SnippetLoadData>
            var trainingDataPath = Path.Combine(Environment.CurrentDirectory, "Data", "recommendation-ratings-train.csv");
            var testDataPath = Path.Combine(Environment.CurrentDirectory, "Data", "recommendation-ratings-test.csv");

            IDataView trainingDataView = mlContext.Data.LoadFromTextFile<MovieRating>(trainingDataPath, hasHeader: true, separatorChar: ',');
            IDataView testDataView = mlContext.Data.LoadFromTextFile<MovieRating>(testDataPath, hasHeader: true, separatorChar: ',');

            return (trainingDataView, testDataView);
            // </SnippetLoadData>
        }

        public static ITransformer BuildAndTrainModel(MLContext mlContext, IDataView trainingDataView)
        {
            // 添加数据转换
            // <SnippetDataTransformations>
            IEstimator<ITransformer> estimator = mlContext.Transforms.Conversion.MapValueToKey(outputColumnName: "userIdEncoded", inputColumnName: "userId")
                .Append(mlContext.Transforms.Conversion.MapValueToKey(outputColumnName: "movieIdEncoded", inputColumnName: "movieId"));
            // </SnippetDataTransformations>

            // 设置算法选项并附加算法
            // <SnippetAddAlgorithm>
            var options = new MatrixFactorizationTrainer.Options
            {
                MatrixColumnIndexColumnName = "userIdEncoded",
                MatrixRowIndexColumnName = "movieIdEncoded",
                LabelColumnName = "Label",
                NumberOfIterations = 20,
                ApproximationRank = 50,
                // 不指定损失函数，将使用默认的平方损失
            };

            var trainerEstimator = estimator.Append(mlContext.Recommendation().Trainers.MatrixFactorization(options));
            // </SnippetAddAlgorithm>

            // <SnippetFitModel>
            Console.WriteLine("=============== Training the model ===============");
            ITransformer model = trainerEstimator.Fit(trainingDataView);

            return model;
            // </SnippetFitModel>
        }




        public static void EvaluateModelAndWriteToExcel(MLContext mlContext, IDataView testDataView, ITransformer model, string outputPath)
        {
            // 在测试数据上评估模型并打印评估指标
            Console.WriteLine("=============== Evaluating the model ===============");
            var prediction = model.Transform(testDataView);

            var metrics = mlContext.Regression.Evaluate(prediction, labelColumnName: "Label", scoreColumnName: "Score");

            Console.WriteLine("Mean Squared Error : " + metrics.MeanSquaredError.ToString());
            Console.WriteLine("Root Mean Squared Error : " + metrics.RootMeanSquaredError.ToString());
            Console.WriteLine("Mean Absolute Error : " + metrics.MeanAbsoluteError.ToString());
            Console.WriteLine("RSquared: " + metrics.RSquared.ToString());

            // 将指标写入到 xlsx 文件
            WriteMetricsToExcel(outputPath, metrics);
        }

        private static void WriteMetricsToExcel(string outputPath, RegressionMetrics metrics)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("EvaluationMetrics");

                // 写入列名
                worksheet.Cells[1, 1].Value = "Metric";
                worksheet.Cells[1, 2].Value = "Value";

                // 写入指标值
                worksheet.Cells[2, 1].Value = "Mean Squared Error";
                worksheet.Cells[2, 2].Value = metrics.MeanSquaredError;

                worksheet.Cells[3, 1].Value = "Root Mean Squared Error";
                worksheet.Cells[3, 2].Value = metrics.RootMeanSquaredError;

                worksheet.Cells[4, 1].Value = "Mean Absolute Error";
                worksheet.Cells[4, 2].Value = metrics.MeanAbsoluteError;

                worksheet.Cells[5, 1].Value = "RSquared";
                worksheet.Cells[5, 2].Value = metrics.RSquared;

                // 保存文件
                package.SaveAs(new FileInfo(outputPath));
            }

            Console.WriteLine($"Metrics written to {outputPath}");
        }



        // 使用模型进行单个预测
        public static void UseModelForSinglePrediction(MLContext mlContext, ITransformer model)
        {
            // <SnippetPredictionEngine>
            Console.WriteLine("=============== Making a prediction ===============");
            var predictionEngine = mlContext.Model.CreatePredictionEngine<MovieRating, MovieRatingPrediction>(model);
            // </SnippetPredictionEngine>

            // 创建测试输入并进行单个预测
            // <SnippetMakeSinglePrediction>
            var testInput = new MovieRating { userId = 9, movieId = 88 };

            var movieRatingPrediction = predictionEngine.Predict(testInput);
            // </SnippetMakeSinglePrediction>

            // <SnippetPrintResults>
            if (Math.Round(movieRatingPrediction.Score, 1) > 3.5)
            {
                Console.WriteLine("Movie " + testInput.movieId + " is recommended for user " + testInput.userId);
            }
            else
            {
                Console.WriteLine("Movie " + testInput.movieId + " is not recommended for user " + testInput.userId);
            }
            // </SnippetPrintResults>
        }

        // 保存模型
        public static void SaveModel(MLContext mlContext, DataViewSchema trainingDataViewSchema, ITransformer model)
        {
            // 将训练好的模型保存到 .zip 文件
            // <SnippetSaveModel>
            var modelPath = Path.Combine(Environment.CurrentDirectory, "Data", "MovieRecommenderModel.zip");

            Console.WriteLine("=============== Saving the model to a file ===============");
            mlContext.Model.Save(model, trainingDataViewSchema, modelPath);
            // </SnippetSaveModel>
        }
    }
}

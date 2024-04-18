using System;
using System.ComponentModel.Design;
using System.Runtime.CompilerServices;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenAI;
using OpenAI.Managers;
using OpenAI.ObjectModels;
using OpenAI.ObjectModels.RequestModels;
using OpenAI.ObjectModels.ResponseModels;
using GPT;

enum Product
{
    EXIT = 0,
    OUTLOOK = 1,
    WORD = 2,
    EXCEL = 3,
    PPT = 4,
    TEST = 5
}

namespace GPT
{
    internal class ChatGptAPI
    {
        internal static string key;
        // API 설정
        OpenAIService openAiService = new OpenAIService(new OpenAiOptions()
        {
            ApiKey = ReadKeyFromConfig()
        });

        // ChatGPT 설정 변수
        ChatCompletionCreateRequest chatRequest;

        // Temperature 설정
        private static double AiTemperatureValue;

        private static string ReadKeyFromConfig()
        {
            try
            {
                using (StreamReader file = File.OpenText("gpt_config.json"))
                {
                    using (JsonTextReader jsonReader = new JsonTextReader(file))
                    {
                        JObject json = (JObject)JToken.ReadFrom(jsonReader);
                        return json["API_KEY"].ToString();
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("Config file not found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading API key: {ex.Message}");
            }

            // 예외가 발생하거나 파일을 찾을 수 없는 경우 null 반환
            return null;
        }

        // 정적 생성자 Temperature 설정 초기화
        static ChatGptAPI()
        {
            AiTemperatureValue = 0.4;
        }

        // ChatGPT 대화의 다양성 설정
        public void SetTemperature(double temperature)
        {
            AiTemperatureValue = temperature;
            ChatGptSetting();
        }

        // ChatGPT Temperature 설정값 넘겨주기
        public double GetTemperature()
        {
            return AiTemperatureValue;
        }

        // 새 대화 생성
        public void ChatGptSetting()
        {
            // ChatGPT 초기 설정
            chatRequest = new ChatCompletionCreateRequest
            {
                Model = Models.Gpt_3_5_Turbo,
                Temperature = (float?)AiTemperatureValue,
                Messages = new List<ChatMessage> { }
            };
        }
        // 질문하기
        public async Task<string> AskQuestion(string prompt)
        {
            // 요청 리퀘스트에 질문 메시지 추가
            ChatMessage chatUserMessage = new ChatMessage("user", prompt);
            chatRequest.Messages.Add(chatUserMessage);

            // 봇 답변 받기
            ChatCompletionCreateResponse chatResult = await openAiService.ChatCompletion.CreateCompletion(chatRequest);

            // 봇 응답이 성공적이면
            if (chatResult.Successful)
            {
                // 봇 답변 내용 저장
                string response = chatResult.Choices.First().Message.Content;

                // 요청 리퀘스트에 답변 메시지 추가, 대화 유지를 위함
                ChatMessage chatBotMessage = new ChatMessage("assistant", response);
                chatRequest.Messages.Add(chatBotMessage);
                return response;
            }
            // 봇 응답이 실패하면
            else
            {
                if (chatResult.Error == null)
                {
                    throw new Exception("Unknown Error");
                }
                return $"Error : {chatResult.Error.Code}: {chatResult.Error.Message}";
            }
        }
    }
}
class Program
{
private static string Get_path()
    {
        Console.Write("Please input target path : ");
        string target_path = Console.ReadLine();

        return target_path;
    }

    private static void Generate_path(string path)
    {
        ChatGptAPI gpt = new ChatGptAPI();

    }

    private static void Init(Product item)
    {
        switch (item)
        {
            case Product.OUTLOOK:
                // TODO : Add Parsing or Conversion API in Outlook
                Console.WriteLine("Selected product : outlook");
                break;
            case Product.WORD:
                // TODO : Add Parsing or Conversion API in Word
                Console.WriteLine("Selected product : word");
                break;
            case Product.EXCEL:
                // TODO : Add Parsing or Conversion API in Excel
                Console.WriteLine("Selected product : excel");
                break;
            case Product.PPT:
                // TODO : Add Parsing or Conversion API in Powerpoint
                Console.WriteLine("Selected product : powerpoint");
                break;
            case Product.TEST:
                break;
            default:
                Console.WriteLine("Exit Program. Good Bye~");
                break;
        }
        string input_path = Get_path();
        Generate_path(input_path);
    }
    public static Product Menu()
    {
        Product selectedItem;
        do
        {
            Console.WriteLine(@"
          ___                       ___                       ___           ___     
         /\  \                     /\  \                     /\__\         /\__\    
        _\:\  \       ___          \:\  \       ___         /:/  /        /:/ _/_   
       /\ \:\  \     /\__\          \:\  \     /\__\       /:/  /        /:/ /\__\  
      _\:\ \:\  \   /:/__/      _____\:\  \   /:/__/      /:/  /  ___   /:/ /:/ _/_ 
     /\ \:\ \:\__\ /::\  \     /::::::::\__\ /::\  \     /:/__/  /\__\ /:/_/:/ /\__\
     \:\ \:\/:/  / \/\:\  \__  \:\~~\~~\/__/ \/\:\  \__  \:\  \ /:/  / \:\/:/ /:/  /
      \:\ \::/  /   ~~\:\/\__\  \:\  \        ~~\:\/\__\  \:\  /:/  /   \::/_/:/  / 
       \:\/:/  /       \::/  /   \:\  \          \::/  /   \:\/:/  /     \:\/:/  /  
        \::/  /        /:/  /     \:\__\         /:/  /     \::/  /       \::/  /   
         \/__/         \/__/       \/__/         \/__/       \/__/         \/__/    
      ");

            Console.WriteLine("Please Select The Object to Analysis");
            Console.WriteLine("== Object List ==");
            Console.WriteLine(@"
      [0] Exit
      [1] Outlook
      [2] Word
      [3] Excel
      [4] Power Point
      ");

            Console.Write("Select item : ");
            if (int.TryParse(Console.ReadLine(), out int item))
            {
                selectedItem = (Product)item;
            }
            else
            {
                Console.WriteLine("Invalid input");
                selectedItem = Product.EXIT;
            }
        } while (selectedItem < Product.EXIT ||  selectedItem > Product.PPT);

        return selectedItem;
    }
    public static void Main(string[] args)
    {
        Product selectedItem = Menu();
        Init(selectedItem);
    }

}
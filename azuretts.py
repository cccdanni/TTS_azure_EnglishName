#!/usr/bin/env python
# coding: utf-8

"""

Speech synthesis for the Microsoft Cognitive Services Speech SDK
Please refer to https://github.com/Azure-Samples/cognitive-services-speech-sdk/blob/master/samples/python/console/speech_synthesis_sample.py
If you wanted to use other functions, i.e., speech_synthesis_to_speaker, 

Copyright (c) Microsoft. All rights reserved.
Licensed under the MIT license. See LICENSE.md file in the project root for full license information.

Author: Chen Danni (dnchen@connect.hku.hk), based on the demo provided by Microsoft
Project: Norms Learning TMR 
History: Create the current file (7/9/2021)

"""

# Import packages 
try:
    import azure.cognitiveservices.speech as speechsdk
except ImportError:
    print("""
    Importing the Speech SDK for Python failed.
    Refer to
    https://docs.microsoft.com/azure/cognitive-services/speech-service/quickstart-text-to-speech-python for
    installation instructions.
    """)
    import sys
    sys.exit(1)
from azure.cognitiveservices.speech import AudioDataStream, SpeechConfig, SpeechSynthesizer, SpeechSynthesisOutputFormat
from azure.cognitiveservices.speech.audio import AudioOutputConfig
import os
import pandas as pd

# Set up the subscription info for the Speech Service:
# *Replace* with your own subscription key and service region (e.g., "westus").
speech_key, service_region = "938024313c36481da2671fe70979078d", "eastus" 
# Create an instance of a speech config with specified subscription key and 
# service region. 
speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=service_region)

# Sets the synthesis language.
# The full list of supported languages can be found here:
# https://docs.microsoft.com/azure/cognitive-services/speech-service/language-support#text-to-speech
language = "en-US"; # *Replace* with your wanted language.
speech_config.speech_synthesis_language = language # Creates a speech synthesizer for the specified language
speech_config.set_speech_synthesis_output_format(SpeechSynthesisOutputFormat["Riff24Khz16BitMonoPcm"]) # 此示例通过对 SpeechConfig 对象设置 SpeechSynthesisOutputFormat 来指定高保真 RIFF 格式 Riff24Khz16BitMonoPcm。 
speech_config.speech_recognition_language = "en"


# Set up parent folder and read word list
work_dir = "C:/Users/Psychology/Desktop/Research/social-norms-learning/Exp2_Norms_TMR/stimulus/cueStimulus"
os.chdir(work_dir)
word_file=('EnglishFirstNameDatabase.xlsx')
word=pd.read_excel(word_file)

for name_nr in range(len(word)):
# for name_nr in range(5):
        word_text = word['First Name'][name_nr]
        word_syllabus = word['Syllabus'][name_nr]
        word_gender = word['Gender'][name_nr]

        fileTitle = 'Names_sounds/' + language + '/' + word_text +'_sy' + str(word_syllabus) + '_' + word_gender + '_' + language + '.wav'
        
        if os.path.isfile(fileTitle) == False :

            # audio_config = AudioOutputConfig(filename=fileTitle)
            # # Create a synthesizer with the given settings. 
            # # Since no explicit audio config is specified, the default speaker will be 
            # # used (make sure the audio settings are correct).
            # synthesizer = SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
            # synthesizer.speak_text_async(word_text)

            synthesizer = SpeechSynthesizer(speech_config=speech_config, audio_config=None)
            result = synthesizer.speak_text_async(word_text).get()
            stream = AudioDataStream(result)
            stream.save_to_wav_file(fileTitle)

            print ('completed: ' + word_text)
        else:
            print ('skip: ' + word_text)


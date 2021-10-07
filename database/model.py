class Question:
	class static_information:
		def __init__(self, question_id, chapter, related_chapter, score, short_answer):
			self.question_id = question_id
			self.chapter = chapter
			self.related_chapter = related_chapter
			self.score = score
			self.short_answer = short_answer

		# Getter of parameters
		def get_chapter(self):
			return self.chapter

		def get_related_chapter(self):
			return self.related_chapter

		def get_score(self):
			return self.get_score

		def get_short_answer(self):
			return self.short_answer

		def get_question_id(self):
			return self.question_id

		# Setter of parameters
		def set_chapter(self, chapter):
			assert isinstance(chapter, str)
			self.chapter = chapter

		def set_related_chapter(self, *related_chapter):
			for i in related_chapter:
				assert isinstance(i, str or int)
			self.related_chapter = list(related_chapter)

		def set_score(self, score):
			assert isinstance(score, int)
			self.score = score

		def set_short_answer(self, short_answer):
			assert isinstance(short_answer, bool)
			self.short_answer = short_answer

		def set_question_id(self, question_id):
			assert isinstance(question_id, str)
			self.question_id = question_id

	class dynamic_information:
		def __init__(self, difficulty, answer_rate, participant, tag, respond):
			self.__difficulty = difficulty
			self.__answer_rate = answer_rate
			self.__participant = participant
			self.__tag = tag
			self.__respond = respond

		# Getter of parameters
		@property
		def difficulty(self):
			return self.__difficulty

		@property
		def answer_rate(self):
			return self.__answer_rate

		@property
		def participant(self):
			return self.__participant

		@property
		def tag(self):
			return self.__tag

		@property
		def respond(self):
			return self.__respond

		# Setter of parameters
		@difficulty.setter
		def difficulty(self, difficulty):
			assert difficulty == ("최상" or "상" or "중" or "하")
			self.__difficulty = difficulty

		@answer_rate.setter
		def answer_rate(self, answer_rate):
			assert isinstance(answer_rate, int or float)
			self.__answer_rate = answer_rate

		@participant.setter
		def participant(self, participant):
			assert isinstance(participant, int)
			self.__participant = participant

		@tag.setter
		def tag(self, tag):
			assert isinstance(self, tag)
			self.__tag = tag

		@respond.setter
		def respond(self, respond):
			self.__respond = respond

# Question-response Pairs